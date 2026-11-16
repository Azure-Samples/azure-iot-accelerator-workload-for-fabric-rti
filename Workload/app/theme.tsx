// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createDarkTheme, createLightTheme, type Theme } from "@fluentui/react-components";
import type { ThemeConfiguration } from "@ms-fabric/workload-client";

export function fabricThemeFromHost(host: ThemeConfiguration): Theme {
    const base = host.colorScheme === "dark" ? createDarkTheme(host.accentColors) : createLightTheme(host.accentColors);
    const overrides = Object.fromEntries(Object.keys(base)
        .filter(key => host.tokens[key] !== undefined)
        .map(key => [key, String(host.tokens[key])]));
    return { ...base, ...overrides };
}