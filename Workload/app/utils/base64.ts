// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

function bytesToBinary(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return binary;
}

export function encodeJsonToBase64(value: unknown): string {
  const json = JSON.stringify(value, null, 2);
  return btoa(bytesToBinary(new TextEncoder().encode(json)));
}

export function decodeJsonFromBase64<T>(payload: string): T {
  const bytes = Uint8Array.from(atob(payload), (character) =>
    character.charCodeAt(0)
  );
  return JSON.parse(new TextDecoder().decode(bytes)) as T;
}
