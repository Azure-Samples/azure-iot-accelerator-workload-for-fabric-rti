// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createBrowserHistory } from "history";
import React from "react";
import { createRoot } from 'react-dom/client';

import { FluentProvider } from "@fluentui/react-components";
import { createWorkloadClient, InitParams, ItemTabActionContext } from '@ms-fabric/workload-client';

import { fabricLightTheme } from "./theme";
import { App } from "./App";
import { callGetItem } from "./controller/ItemCRUDController"

export async function initialize(params: InitParams) {
    const workloadClient = createWorkloadClient();
    const history = createBrowserHistory();

    workloadClient.navigation.onNavigate((route) => {
        history.replace(route.targetUrl);
    });
    workloadClient.action.onAction(async function ({ action, data }) {
        const { id } = data as ItemTabActionContext;
        switch (action) {
            case 'item.tab.onInit':
                try {
                    const itemResult = await callGetItem(workloadClient, id);
                    if (itemResult?.item?.displayName) {
                        return { title: itemResult.item.displayName };
                    } else {
                        console.warn(`Item not found or missing displayName for ID: ${id}`);
                        return { title: 'Untitled Item' }; // Provide a default title
                    }
                } catch (error) {
                    console.error(
                        `Error loading the Item (object ID:${id})`,
                        error
                    );
                    return {};
                }
            case 'item.tab.canDeactivate':
                return { canDeactivate: true };
            case 'item.tab.onDeactivate':
                return {};
            case 'item.tab.canDestroy':
                return { canDestroy: true };
            case 'item.tab.onDestroy':
                return {};
            case 'item.tab.onDelete':
                return {};
            default:
                throw new Error('Unknown action received');
        }
    });
    
    const rootElement = document.getElementById('root');
    if (!rootElement) {
        console.error('Root element not found');
        return;
    }

    const root = createRoot(rootElement);
    root.render(
        <FluentProvider theme={fabricLightTheme}>
            <App history={history} workloadClient={workloadClient} />
        </FluentProvider>
    );
}
