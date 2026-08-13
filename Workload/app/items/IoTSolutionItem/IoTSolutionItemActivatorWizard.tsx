import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardControl, WizardStep, WizardStepProps } from "../../components/Wizard";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { ActivatorOverviewStep } from "./steps/ActivatorOverviewStep";
import { ActivatorSelectTableStep } from "./steps/ActivatorSelectTableStep";
import { ActivatorWorkspaceIdentityStep } from "./steps/ActivatorWorkspaceIdentityStep";
import { ActivatorCreateResourcesStep } from "./steps/ActivatorCreateResourcesStep";
import { ActivatorSummaryStep } from "./steps/ActivatorSummaryStep";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemActivatorWizardProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onWizardComplete: (context: Record<string, any>) => void;
  onBack?: () => void;
}

/**
 * Activator wizard with outcome-focused overview and completion pages.
 */
export function IoTSolutionItemActivatorWizard({
  workloadClient,
  item,
  onWizardComplete,
  onBack,
}: IoTSolutionItemActivatorWizardProps) {
  const workspaceId = item?.workspaceId || "";
  const itemId = item?.id || "default";
  const SelectTableStepWrapper = (props: WizardStepProps) => (
    <ActivatorSelectTableStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
    />
  );

  const WorkspaceIdentityStepWrapper = (props: WizardStepProps) => (
    <ActivatorWorkspaceIdentityStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
    />
  );

  const CreateResourcesStepWrapper = (props: WizardStepProps) => (
    <ActivatorCreateResourcesStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
    />
  );

  const SummaryStepWrapper = (props: WizardStepProps) => (
    <ActivatorSummaryStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
    />
  );

  const steps: WizardStep[] = [
    {
      id: "overview",
      title: "Overview",
      description: "Monitor device conditions and automate responses",
      component: ActivatorOverviewStep,
      validate: () => true,
      numbered: false,
    },
    {
      id: "select-table",
      title: "Monitoring Data",
      description: "Select the modeled device data to evaluate",
      component: SelectTableStepWrapper,
      validate: (context) => !!context.actTableName,
    },
    {
      id: "workspace-identity",
      title: "Workspace Identity",
      description: "Enable secure Fabric resource access",
      component: WorkspaceIdentityStepWrapper,
      validate: (context) => !!context.actWorkspaceIdentityReady,
    },
    {
      id: "create-resources",
      title: "Activator Resources",
      description: "Create the live data feed and Activator",
      component: CreateResourcesStepWrapper,
      validate: (context) => !!context.activatorCreated,
      isNavigationDisabled: (context) => !!context.actResourcesCreating,
    },
    {
      id: "summary",
      title: "Ready to Configure",
      description: "Review the monitoring flow and create rules",
      component: SummaryStepWrapper,
      numbered: false,
      validate: (context) => !!context.activatorCreated,
    },
  ];

  return (
    <WizardControl
      title="Set Up Activator Rules"
      steps={steps}
      onComplete={onWizardComplete}
      canFinish={(_stepId, context) => !!context.activatorCreated}
      showNavigation={true}
      allowStepNavigation={false}
      navigationLabels={{
        next: "Next",
        previous: "Previous",
        complete: "Finish and Close",
        cancel: "Cancel",
      }}
      onFirstStepBack={onBack}
      persistKey={`iot-wizard-v4-activator-${itemId}`}
    />
  );
}
