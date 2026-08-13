import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { WizardControl, WizardStep, WizardStepProps } from "../../components/Wizard";
import { IngestionOverviewStep } from "./steps/IngestionOverviewStep";
import { IoTHubStep } from "./steps/IoTHubStep";
import { EventhouseStep } from "./steps/EventhouseStep";
import { TableSetupStep } from "./steps/TableSetupStep";
import { EventstreamStep } from "./steps/EventstreamStep";
import { RoutingStep } from "./steps/RoutingStep";
import { IngestionSummaryStep } from "./steps/IngestionSummaryStep";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemDefaultViewProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onWizardComplete: (context: Record<string, any>) => void;
  onBack?: () => void;
}

/**
 * Default view for the IoT Solution item.
 * Renders the setup wizard with progressive steps.
 */
export function IoTSolutionItemDefaultView({
  workloadClient,
  item,
  onWizardComplete,
  onBack,
}: IoTSolutionItemDefaultViewProps) {

  const workspaceId = item?.workspaceId || "";

  // Wrap IoTHubStep to inject workloadClient
  const IoTHubStepWrapper = (props: WizardStepProps) => (
    <IoTHubStep {...props} workloadClient={workloadClient} />
  );

  // Wrap EventhouseStep to inject workloadClient and workspaceId
  const EventhouseStepWrapper = (props: WizardStepProps) => (
    <EventhouseStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  // Wrap TableSetupStep to inject workloadClient and workspaceId
  const TableSetupStepWrapper = (props: WizardStepProps) => (
    <TableSetupStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  // Wrap EventstreamStep to inject workloadClient and workspaceId
  const EventstreamStepWrapper = (props: WizardStepProps) => (
    <EventstreamStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  // Wrap RoutingStep to inject workloadClient and workspaceId
  const RoutingStepWrapper = (props: WizardStepProps) => (
    <RoutingStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  const steps: WizardStep[] = [
    {
      id: "overview",
      title: "Overview",
      description: "Connect IoT Hub data to Fabric analytics",
      component: IngestionOverviewStep,
      validate: () => true,
      numbered: false,
    },
    {
      id: "iot-hub",
      title: "IoT Hub",
      description: "Select your Azure IoT Hub",
      component: IoTHubStepWrapper,
      validate: async (context: Record<string, any>) => {
        return !!context.iotHubValidated;
      },
    },
    {
      id: "eventhouse",
      title: "Eventhouse",
      description: "Select an Eventhouse and database",
      component: EventhouseStepWrapper,
      validate: async (context: Record<string, any>) => {
        return !!context.eventhouseValidated;
      },
    },
    {
      id: "tables",
      title: "Raw Data Tables",
      description: "Create KQL tables for telemetry and properties",
      component: TableSetupStepWrapper,
      validate: async (context: Record<string, any>) => {
        return !!context.tablesCreated;
      },
      isNavigationDisabled: (context) => !!context.tableSetupCreating,
    },
    {
      id: "eventstreams",
      title: "Eventstreams",
      description: "Create Eventstreams with custom endpoints",
      component: EventstreamStepWrapper,
      validate: async (context: Record<string, any>) => {
        return !!context.eventstreamsCreated;
      },
      isNavigationDisabled: (context) => !!context.eventstreamsCreating,
    },
    {
      id: "routing",
      title: "IoT Hub Routing",
      description: "Configure message routing to Eventstreams",
      component: RoutingStepWrapper,
      validate: async (context: Record<string, any>) => {
        return !!context.routingConfigured;
      },
    },
    {
      id: "summary",
      title: "Ready to Use",
      description: "Review the architecture and next steps",
      component: IngestionSummaryStep,
      numbered: false,
      validate: async (context: Record<string, any>) => {
        return !!context.routingConfigured;
      },
    },
  ];

  // Always start with a blank context — the wizard creates fresh resources each run
  const initialContext: Record<string, any> = {};

  const itemId = item?.id || "default";

  return (
    <WizardControl
      title="Ingest device data from your IoT Hub"
      steps={steps}
      initialContext={initialContext}
      onComplete={onWizardComplete}
      canFinish={(_stepId, context) =>
        !!context.iotHubValidated &&
        !!context.eventhouseValidated &&
        !!context.tablesCreated &&
        !!context.eventstreamsCreated &&
        !!context.routingConfigured
      }
      showNavigation={true}
      allowStepNavigation={true}
      navigationLabels={{
        next: "Next",
        previous: "Previous",
        complete: "Finish and Close",
        cancel: "Cancel",
      }}
      onFirstStepBack={onBack}
      persistKey={`iot-wizard-v4-ingestion-${itemId}`}
    />
  );
}
