// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { JsonObject, JsonValue } from "./ModelAuthoring.ts";

export const DOCSTRING_LIMIT = 1000;
export const docstringJson = (value: JsonValue): string => JSON.stringify(value).replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
export type DocstringPart = ({ text: string } | { key: string; value: JsonValue } |
  { block: string; children: readonly DocstringPart[] }) & { required?: boolean; shorten?: boolean };
export function isDocstringValue(value: unknown): value is JsonValue {
  return value === null || typeof value === "string" || typeof value === "boolean" ||
    typeof value === "number" && Number.isFinite(value) ||
    Array.isArray(value) && value.every(isDocstringValue) ||
    typeof value === "object" && value !== null && Object.values(value).every(isDocstringValue);
}
const render = (part: DocstringPart, depth = 0): string => {
  const indent = "\t".repeat(depth);
  if ("block" in part) return `${indent}${part.block}:\n${part.children.map(child => render(child, depth + 1)).join("\n")}`;
  return indent + ("text" in part ? part.text : `${part.key}: ${docstringJson(part.value)}`);
};

/** Bound stored text, not its KQL/JSON transport representation. Never cut control lines or JSON values. */
export function emitDocstring(parts: readonly DocstringPart[]): string {
  const validate = (values: readonly DocstringPart[], depth: number) => {
    if (depth > 48) throw new Error("Docstring metadata exceeds the supported nesting depth.");
    for (const part of values) {
      if ("text" in part && /[\r\n\u2028\u2029]/.test(part.text) || "key" in part && /[:\r\n\u2028\u2029]/.test(part.key) ||
          "block" in part && /[\r\n\u2028\u2029]/.test(part.block)) {
        throw new Error("Docstring lines must not contain embedded control-line delimiters.");
      }
      if ("key" in part && !isDocstringValue(part.value)) throw new Error("Docstring facts must be finite JSON values.");
      if ("block" in part) validate(part.children, depth + 1);
    }
  };
  validate(parts, 0);
  const full = parts.map(part => render(part)).join("\n");
  if (full.length <= DOCSTRING_LIMIT) return full;
  const marker = "truncated: true";
  const required = parts.filter(part => part.required).map(part => render(part));
  if (required.join("\n").length > DOCSTRING_LIMIT) {
    throw new Error("Required docstring identity metadata exceeds the 1000-character limit. Use a source/model identity combination that fits; identities were not shortened.");
  }
  if ([...required, marker].join("\n").length > DOCSTRING_LIMIT && parts.some(part => !part.required)) {
    throw new Error("Required docstring identity metadata leaves no room to disclose bounded human facts within 1000 characters.");
  }
  const fit = (part: DocstringPart, budget: number, depth: number): string | undefined => {
    const line = render(part, depth);
    if (line.length <= budget) return line;
    if ("block" in part) {
      const header = `${"\t".repeat(depth)}${part.block}:`;
      const children = fitParts(part.children, budget - header.length - 1, depth + 1);
      return children.length ? `${header}\n${children.join("\n")}` : undefined;
    }
    const value = "text" in part ? part.text : part.value;
    if (typeof value !== "string" || part.required || "key" in part && !part.shorten) return undefined;
    const points = Array.from(value);
    let low = 0, high = points.length, fitted: string | undefined;
    while (low <= high) {
      const count = Math.floor((low + high) / 2), shortened = `${points.slice(0, count).join("")}...`;
      const candidate = "\t".repeat(depth) + ("text" in part ? shortened : `${part.key}: ${docstringJson(shortened)}`);
      if (candidate.length <= budget) { fitted = candidate; low = count + 1; } else high = count - 1;
    }
    return fitted;
  };
  const fitParts = (values: readonly DocstringPart[], budget: number, depth: number): string[] => {
    const mandatory = values.filter(part => part.required).map(part => render(part, depth));
    if (mandatory.join("\n").length > budget) return [];
    const lines: string[] = [];
    let requiredLeft = mandatory.length;
    for (const part of values) {
      if (part.required) { lines.push(render(part, depth)); requiredLeft--; continue; }
      const remaining = mandatory.slice(mandatory.length - requiredLeft);
      const rest = [...lines, ...remaining].join("\n").length;
      const available = budget - rest - (lines.length || remaining.length ? 1 : 0);
      const fitted = fit(part, available, depth);
      if (fitted) lines.push(fitted);
    }
    return lines;
  };
  const lines = fitParts(parts, DOCSTRING_LIMIT - marker.length - 1, 0);
  const result = [...lines, marker].join("\n");
  if (result.length > DOCSTRING_LIMIT) throw new Error("The generated docstring could not be bounded safely.");
  return result;
}

export function docstringControl(text: string): DocstringPart {
  if (/[\r\n\u2028\u2029]/.test(text)) throw new Error("Docstring identity metadata cannot contain line breaks.");
  return { text, required: true };
}

/** JSON is budgeted as complete values before it reaches the common stored-text boundary. */
export function emitJsonDocstring(prefix: string, value: JsonObject,
  surrounding?: { before: readonly string[]; after: readonly string[]; requiredKeys: readonly string[] }): string {
  const validate = (item: JsonValue, depth: number) => {
    if (depth > 48 || item === undefined || typeof item === "number" && !Number.isFinite(item)) {
      throw new Error("Docstring semantic JSON must contain finite values within the supported nesting depth.");
    }
    if (Array.isArray(item)) item.forEach(child => validate(child, depth + 1));
    else if (item !== null && typeof item === "object") Object.values(item).forEach(child => validate(child, depth + 1));
  };
  validate(value, 0);
  const before = surrounding?.before || [], after = surrounding?.after || [];
  const payload = (item: JsonObject) => [...before, prefix + docstringJson(item), ...after].join("\n");
  const emit = (item: JsonObject) => emitDocstring([
    ...before.map(docstringControl), docstringControl(prefix + docstringJson(item)), ...after.map(docstringControl),
  ]);
  const subject = surrounding ? "model" : "column";
  if (payload(value).length <= DOCSTRING_LIMIT) return emit(value);
  const human = new Set(["title", "titles", "displayName", "description", "descriptions", "comment"]);
  const required = (key: string, item: JsonValue) => (surrounding?.requiredKeys || ["name", "component", "role", "@id"]).includes(key) ||
    typeof item === "boolean" || item === null || item === "" ||
    typeof item === "number" && docstringJson(item).length <= 20 ||
    ["unit", "type", "schema"].includes(key) && typeof item === "string" && docstringJson(item).length <= 100 ||
    ["const", "default"].includes(key) && typeof item === "string" && docstringJson(item).length <= 20;
  const entries = Object.entries(value);
  const result: JsonObject = Object.fromEntries(entries.filter(([key, item]) => required(key, item)));
  if (payload(result).length > DOCSTRING_LIMIT) throw new Error(`Required ${subject} semantic identity/role metadata exceeds the 1000-character limit; identities were not shortened.`);
  result.truncated = true;
  const fits = (item: JsonObject) => payload(item).length <= DOCSTRING_LIMIT;
  if (!fits(result)) throw new Error(`Required ${subject} semantic metadata leaves no room for an honest truncation signal within 1000 characters.`);
  for (const [key, item] of entries.filter(([key, item]) => !required(key, item) && !human.has(key))) {
    const candidate = { ...result, [key]: item };
    if (fits(candidate)) Object.defineProperty(result, key, { value: item, enumerable: true, configurable: true, writable: true });
  }
  for (const [key, item] of entries.filter(([key, item]) => !required(key, item) && human.has(key))) {
    if (fits({ ...result, [key]: item })) { result[key] = item; continue; }
    const texts = typeof item === "string" ? [item] : item !== null && !Array.isArray(item) && typeof item === "object" &&
      Object.values(item).every(value => typeof value === "string") ? Object.values(item) : [];
    if (!texts.length) continue;
    const shorten = (text: JsonValue, count: number): JsonValue => {
      if (typeof text !== "string") return text;
      const points = Array.from(text);
      return points.length > count ? points.slice(0, count).join("") + "..." : text;
    };
    let low = 0, high = texts.reduce<number>((maximum, text) => Math.max(maximum, typeof text === "string" ? Array.from(text).length : 0), 0);
    let fitted: JsonValue | undefined;
    while (low <= high) {
      const count = Math.floor((low + high) / 2);
      const candidate = typeof item === "string" ? shorten(item, count)
        : Object.fromEntries(Object.entries(item).map(([locale, text]) => [locale, shorten(text, count)]));
      if (fits({ ...result, [key]: candidate })) { fitted = candidate; low = count + 1; } else high = count - 1;
    }
    if (fitted !== undefined) result[key] = fitted;
  }
  return emit(result);
}

export function parseDocstringJson(text: string, subject = "column"): JsonObject {
  let value: unknown;
  try { value = JSON.parse(text); } catch { throw new Error(`Malformed ${subject} semantic JSON.`); }
  if (value === null || typeof value !== "object" || Array.isArray(value)) throw new Error(`${subject} semantic JSON must be an object.`);
  const stack: { object: boolean; keys: Set<string>; key: boolean }[] = [];
  for (const token of text.match(/"(?:[^"\\]|\\.)*"|[{}\[\],:]/g) || []) {
    if (token === "{" || token === "[") {
      stack.push({ object: token === "{", keys: new Set<string>(), key: token === "{" });
      if (stack.length > 48) throw new Error(`${subject} semantic JSON exceeds the supported nesting depth.`);
    } else if (token === "}" || token === "]") stack.pop();
    else if (token === "," && stack[stack.length - 1]?.object) stack[stack.length - 1].key = true;
    else if (token.startsWith('"') && stack[stack.length - 1]?.key) {
      const frame = stack[stack.length - 1], key: string = JSON.parse(token);
      if (frame.keys.has(key)) throw new Error(`Duplicate ${subject} semantic JSON key.`);
      frame.keys.add(key); frame.key = false;
    }
  }
  if (!isDocstringValue(value) || Array.isArray(value)) throw new Error(`${subject} semantic JSON contains a non-finite value.`);
  return value;
}
