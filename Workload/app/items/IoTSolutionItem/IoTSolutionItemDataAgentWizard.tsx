import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardControl, WizardStep, WizardStepProps } from "../../components/Wizard";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { DataAgentOverviewStep } from "./steps/DataAgentOverviewStep";
import { DataAgentSetupStep } from "./steps/DataAgentSetupStep";
import { DataAgentSummaryStep } from "./steps/DataAgentSummaryStep";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemDataAgentWizardProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onWizardComplete: (context: Record<string, any>) => void;
  onBack?: () => void;
}

/**
 * Data Agent wizard with separate data-source, configuration, and completion
 * experiences.
 */
export function IoTSolutionItemDataAgentWizard({
  workloadClient,
  item,
  onWizardComplete,
  onBack,
}: IoTSolutionItemDataAgentWizardProps) {
  const workspaceId = item?.workspaceId || "";
  const itemId = item?.id || "default";
  const DataSourceStepWrapper = (props: WizardStepProps) => (
    <DataAgentSetupStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
      phase="source"
    />
  );

  const ConfigureAgentStepWrapper = (props: WizardStepProps) => (
    <DataAgentSetupStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
      phase="configure"
    />
  );

  const steps: WizardStep[] = [
    {
      id: "overview",
      title: "Overview",
      description: "Ask questions grounded in device operations",
      component: DataAgentOverviewStep,
      validate: () => true,
      numbered: false,
    },
    {
      id: "data-source",
      title: "Modeled Data Source",
      description: "Select the device data that grounds the agent",
      component: DataSourceStepWrapper,
      validate: (context) => !!context.agentSourceSelected,
    },
    {
      id: "configure-agent",
      title: "Configure Data Agent",
      description: "Review instructions and create the agent",
      component: ConfigureAgentStepWrapper,
      validate: (context) => !!context.agentCreated,
      isNavigationDisabled: (context) => !!context.agentCreating,
    },
    {
      id: "summary",
      title: "Ready to Use",
      description: "Review the agent and example questions",
      component: DataAgentSummaryStep,
      numbered: false,
      validate: (context) => !!context.agentCreated,
    },
  ];

  return (
    <WizardControl
      title="Set up Data Agent"
      steps={steps}
      onComplete={onWizardComplete}
      canFinish={(_stepId, context) => !!context.agentCreated}
      showNavigation={true}
      allowStepNavigation={false}
      navigationLabels={{
        next: "Next",
        previous: "Previous",
        complete: "Finish and Close",
        cancel: "Cancel",
      }}
      onFirstStepBack={onBack}
      persistKey={`iot-wizard-v4-dataagent-${itemId}`}
    />
  );
}
