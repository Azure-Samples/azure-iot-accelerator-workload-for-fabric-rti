type ErrorRecord = Record<string, unknown>;

function asRecord(value: unknown): ErrorRecord | undefined {
  return typeof value === "object" && value !== null
    ? value as ErrorRecord
    : undefined;
}

function firstString(...values: unknown[]): string {
  return values.find((value): value is string => typeof value === "string" && value.trim().length > 0) || "";
}

export function parseFabricErrorPayload(text: string): unknown {
  if (!text.trim()) return {};

  try {
    return JSON.parse(text);
  } catch {
    return { message: text };
  }
}

export function hasFabricErrorCode(payload: unknown, expectedCode: string): boolean {
  const response = asRecord(payload);
  const nestedError = asRecord(response?.error);
  return firstString(
    response?.errorCode,
    response?.code,
    nestedError?.errorCode,
    nestedError?.code
  ) === expectedCode;
}

export function getFabricCreationErrorMessage(
  status: number,
  payload: unknown,
  resourceType: string,
  displayName: string
): string {
  const response = asRecord(payload);
  const nestedError = asRecord(response?.error);
  const code = firstString(
    response?.errorCode,
    response?.code,
    nestedError?.errorCode,
    nestedError?.code
  );
  const serviceMessage = firstString(response?.message, nestedError?.message);
  const resourceLabel = resourceType.toLowerCase();

  if (code === "ItemDisplayNameAlreadyInUse") {
    return `A ${resourceLabel} named "${displayName}" already exists in this workspace. Choose a different name and try again.`;
  }

  if (code === "DuplicateConnectionName") {
    return `A connection named "${displayName}" already exists and could not be recovered. Remove the unused connection in Fabric and try again.`;
  }

  // Some asynchronous operation responses return a generic Conflict code.
  if (/\b(already exists|already in use)\b/i.test(serviceMessage)) {
    return `A ${resourceLabel} named "${displayName}" already exists in this workspace. Choose a different name and try again.`;
  }

  if (status === 401 || status === 403) {
    return `Unable to create the ${resourceLabel}. Verify that you have permission to create it in this workspace.`;
  }

  if (status === 429) {
    return `Fabric is temporarily limiting ${resourceLabel} creation requests. Wait a moment and try again.`;
  }

  if (serviceMessage) {
    return `Fabric could not create the ${resourceLabel}: ${serviceMessage}`;
  }

  return `Fabric could not create the ${resourceLabel} (HTTP ${status}). Try again.`;
}
