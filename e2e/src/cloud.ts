import { AzureCliCredential } from "@azure/identity";
import { existsSync } from "node:fs";
import { delimiter, isAbsolute, join } from "node:path";
import type { RunConfig } from "./config.js";

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

export function object(value: unknown, label: string): Record<string, unknown> {
  if (!record(value)) {
    throw new Error(`[service-shape] ${label} must be an object.`);
  }
  return value;
}

export function text(value: unknown, label: string): string {
  if (typeof value !== "string" || !value) throw new Error(`[service-shape] ${label} is missing.`);
  return value;
}

export function safeServiceMessage(value: unknown): string {
  if (typeof value !== "string") return "";
  return redactSensitiveText(value).slice(0, 2000);
}

export function redactSensitiveText(value: string): string {
  return value
    .replace(/https?:\/\/[^\s"'<>\\]+/g, url => url.replace(/([?&])([^=&#]+)=([^&#]*)/g,
      (parameter, separator: string, key: string) => key.toLowerCase() === "api-version" ? parameter : `${separator}${key}=[REDACTED]`))
    .replace(/\bBearer\s+\S+/gi, "Bearer [REDACTED]")
    .replace(/\bSharedAccessSignature\s+[^\s"']+/gi, "SharedAccessSignature [REDACTED]")
    .replace(/((?:SharedAccessKey|primaryKey|secondaryKey|access_token|refresh_token|client_secret)["']?\s*[:=]\s*["']?)[^;"'\s,}]+/gi, "$1[REDACTED]")
    .replace(/\beyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b/g, "[REDACTED JWT]");
}

export class CloudError extends Error {
  constructor(readonly status: number, readonly code: string, method: string, url: URL, requestId: string | null, message?: unknown) {
    super(`[${status === 401 || status === 403 ? "permissions" : "service"}] ${method} ${url.origin}${url.pathname}: HTTP ${status}, ${code}${requestId ? `, request ${requestId}` : ""}${message ? `: ${safeServiceMessage(message)}` : ""}`);
  }
}

export class Cloud {
  private readonly credential: AzureCliCredential;
  private readonly tokens = new Map<string, { token: string; expiresOnTimestamp: number }>();

  constructor(readonly config: RunConfig) {
    this.credential = new AzureCliCredential({ subscription: config.subscriptionId });
  }

  async token(resource: string): Promise<string> {
    const cached = this.tokens.get(resource);
    if (cached && cached.expiresOnTimestamp > Date.now() + 60_000) return cached.token;
    const token = await this.credential.getToken(`${resource.replace(/\/$/, "")}/.default`).catch(failure => {
      try {
        const pathKeys = Object.keys(process.env).filter(key => key.toLowerCase() === "path");
        const paths = pathKeys.flatMap(key => (process.env[key] ?? "").split(delimiter))
          .map(value => value.trim().replace(/^"(.*)"$/, "$1")).filter(isAbsolute);
        const systemRoot = process.env.SystemRoot ?? process.env.SYSTEMROOT;
        const comSpec = process.env.ComSpec ?? process.env.COMSPEC;
        console.error(`[authentication-process] ${JSON.stringify({
          platform: process.platform, cwdExists: existsSync(process.cwd()), executableExists: existsSync(process.execPath),
          pathKeys, pathEntries: paths.length,
          azureCliFileOnPath: paths.some(path => (process.platform === "win32" ? ["az.exe", "az.cmd", "az.bat"] : ["az"])
            .some(name => existsSync(join(path, name)))),
          systemRootExists: !!systemRoot && existsSync(systemRoot), comSpecExists: !!comSpec && existsSync(comSpec),
        })}`);
      } catch (diagnosticFailure) {
        console.error(`[authentication-process] Inspection failed: ${redactSensitiveText(diagnosticFailure instanceof Error ? diagnosticFailure.message : String(diagnosticFailure))}`);
      }
      throw failure;
    });
    const payload = token.token.split(".")[1];
    if (!payload) throw new Error("[authentication] The CLI token is not a JWT.");
    const claims = object(JSON.parse(Buffer.from(payload, "base64url").toString("utf8")), "token claims");
    if (claims.tid !== this.config.tenantId) throw new Error("[authentication] Azure CLI token tenant differs from the requested tenant.");
    this.tokens.set(resource, token);
    return token.token;
  }

  async request(url: string, resource: string, method = "GET", body?: unknown, missingAllowed = false, ifMatch?: string, ifNoneMatch?: string): Promise<unknown> {
    return (await this.send(url, resource, method, body, missingAllowed, ifMatch, ifNoneMatch)).body;
  }

  async send(url: string, resource: string, method = "GET", body?: unknown, missingAllowed = false, ifMatch?: string, ifNoneMatch?: string): Promise<{ body: unknown; status: number; headers: Headers }> {
    const target = new URL(url);
    if (target.protocol !== "https:" || target.username || target.password) throw new Error("[configuration] Cloud requests require a credential-free HTTPS URL.");
    const token = await this.token(resource);
    const response = await fetch(target, {
      method,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Connection: "close", ...(ifMatch ? { "If-Match": ifMatch } : {}), ...(ifNoneMatch ? { "If-None-Match": ifNoneMatch } : {}) },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      signal: AbortSignal.timeout(60_000),
      redirect: "error",
    });
    if (response.status === 404 && missingAllowed) {
      await response.body?.cancel();
      return { body: undefined, status: 404, headers: response.headers };
    }
    const raw = await response.text();
    let result: unknown;
    if (raw) {
      try {
        result = JSON.parse(raw);
      } catch (error) {
        if (!(error instanceof SyntaxError)) throw error;
        throw new Error(`[service-shape] ${method} ${target.origin}${target.pathname} returned non-JSON content (HTTP ${response.status}).`);
      }
    }
    if (!response.ok) {
      const error = result === undefined ? {} : object(result, "error response");
      const nested = error.error === undefined ? error : object(error.error, "error details");
      const code = nested.code ?? nested.errorCode ?? nested.Code ?? nested.ErrorCode;
      throw new CloudError(response.status,
        typeof code === "string" || typeof code === "number" ? String(code) : "Unspecified",
        method, target, response.headers.get("x-ms-request-id") ?? response.headers.get("requestId"),
        nested.message ?? nested.Message);
    }
    return { body: result, status: response.status, headers: response.headers };
  }

  arm(path: string, version: string, method = "GET", body?: unknown, missingAllowed = false, ifMatch?: string) {
    if (!path.startsWith(`/subscriptions/${this.config.subscriptionId}/`) && path !== `/subscriptions/${this.config.subscriptionId}`) {
      throw new Error("[ownership] ARM request is outside the configured subscription.");
    }
    return this.request(`https://management.azure.com${path}?api-version=${version}`, "https://management.azure.com", method, body, missingAllowed, ifMatch);
  }

  fabric(path: string, method = "GET", body?: unknown) {
    if (!path.startsWith(`/workspaces/${this.config.workspaceId}/`) && !/^\/operations\/[a-f0-9-]+(?:\/result)?$/i.test(path)) {
      throw new Error("[ownership] Fabric request is outside the supplied workspace or operation.");
    }
    return this.request(`https://api.fabric.microsoft.com/v1${path}`, "https://api.fabric.microsoft.com", method, body);
  }
}
