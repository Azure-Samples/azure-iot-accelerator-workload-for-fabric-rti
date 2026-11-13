// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useCallback, useEffect, useRef, useState } from "react";
import { Button, FluentProvider, type Theme } from "@fluentui/react-components";
import type { ThemeConfiguration, WorkloadClientAPI } from "@ms-fabric/workload-client";
import { fabricThemeFromHost } from "./theme";

export function FabricHostThemeProvider({ workloadClient, children }: {
    workloadClient: WorkloadClientAPI;
    children: React.ReactNode;
}) {
    const [theme, setTheme] = useState<Theme>();
    const [scheme, setScheme] = useState("normal");
    const [error, setError] = useState("");
    const revision = useRef(0);
    const active = useRef(true);
    const apply = useCallback((configuration: ThemeConfiguration) => {
        setTheme(fabricThemeFromHost(configuration));
        setScheme(configuration.colorScheme);
        setError("");
    }, []);
    const load = useCallback(async () => {
        const version = revision.current;
        setError("");
        try {
            const configuration = await workloadClient.theme.get();
            if (active.current && revision.current === version) apply(configuration);
        } catch (failure: unknown) {
            if (active.current && revision.current === version) {
                console.error("Unable to load the Fabric host theme.", failure);
                setError("Unable to load the Fabric theme. Try again.");
            }
        }
    }, [workloadClient, apply]);
    useEffect(() => {
        active.current = true;
        let subscribed = true;
        workloadClient.theme.onChange(configuration => {
            if (!subscribed) return;
            revision.current++;
            apply(configuration);
        });
        void load();
        return () => { subscribed = false; active.current = false; };
    }, [workloadClient, apply, load]);
    const problem = error && <div role="alert">{error} <Button onClick={load}>Try again</Button></div>;
    // Keep required app initialization independent of the asynchronous theme request.
    return <>
        {!theme && (problem || <div role="status">Loading Fabric theme...</div>)}
        <FluentProvider theme={theme} aria-hidden={theme ? undefined : true}
            style={{ colorScheme: scheme, visibility: theme ? undefined : "hidden" }}>
            {theme && problem}{children}
        </FluentProvider>
    </>;
}
