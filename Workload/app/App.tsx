import React from "react";
import { Route, Router, Switch } from "react-router-dom";
import { History } from "history";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { IoTSolutionItemEditor } from "./items/IoTSolutionItem";

/*
    Add your Item Editor in the Route section of the App function below
*/

interface AppProps {
    history: History;
    workloadClient: WorkloadClientAPI;
}

export interface PageProps {
    workloadClient: WorkloadClientAPI;
    history?: History
}

export interface ContextProps {
    itemObjectId?: string;
    workspaceObjectId?: string
    source?: string;
}

export function App({ history, workloadClient }: AppProps) {
    return <Router history={history}>
        <Switch>
            {/* Routing for the IoT Solution Item Editor */}
            <Route path="/IoTSolutionItem-editor/:itemObjectId">
                <IoTSolutionItemEditor
                    workloadClient={workloadClient} data-testid="IoTSolutionItem-editor" />
            </Route>
        </Switch>
    </Router>;
}