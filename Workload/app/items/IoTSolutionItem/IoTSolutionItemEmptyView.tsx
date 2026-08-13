import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { ItemEditorEmptyView, EmptyStateTask } from "../../components/ItemEditor";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemEmptyViewProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onStartSetup: () => void;
}

/**
 * Empty state view shown when a new IoT Solution item is created.
 * Provides a single call-to-action to begin the setup wizard.
 */
export function IoTSolutionItemEmptyView({
  onStartSetup,
}: IoTSolutionItemEmptyViewProps) {
  const tasks: EmptyStateTask[] = [
    {
      id: "start-setup",
      label: "Start Setup Wizard",
      icon: undefined,
      description: "Configure your IoT Hub connection, set up data ingestion, and create dashboards.",
      onClick: onStartSetup,
    },
  ];

  return (
    <ItemEditorEmptyView
      title="Welcome to IoT Solution"
      description="This wizard will guide you through connecting your Azure IoT Hub to Microsoft Fabric, setting up telemetry ingestion, and creating real-time dashboards."
      imageSrc="/assets/items/IoTSolutionItem/EditorEmpty.svg"
      imageAlt="IoT Solution setup illustration"
      tasks={tasks}
    />
  );
}
