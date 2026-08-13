import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardControl, WizardStep, WizardStepProps } from "../../components/Wizard";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { ModelOverviewStep } from "./steps/ModelOverviewStep";
import { DeviceModelStep } from "./steps/DeviceModelStep";
import { ModelCoverageStep } from "./steps/ModelCoverageStep";
import { ModeledDataSetupStep } from "./steps/ModeledDataSetupStep";
import { ModelSummaryStep } from "./steps/ModelSummaryStep";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemModelWizardProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onWizardComplete: (context: Record<string, any>) => void;
  onBack?: () => void;
}

/**
 * Device Model wizard with outcome-focused overview and completion pages.
 */
export function IoTSolutionItemModelWizard({
  workloadClient,
  item,
  onWizardComplete,
  onBack,
}: IoTSolutionItemModelWizardProps) {
  const workspaceId = item?.workspaceId || "";
  const itemId = item?.id || "default";

  const DeviceModelStepWrapper = (props: WizardStepProps) => (
    <DeviceModelStep {...props} />
  );

  const RawDataStepWrapper = (props: WizardStepProps) => (
    <ModeledDataSetupStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
      phase="source"
    />
  );

  const CoverageStepWrapper = (props: WizardStepProps) => (
    <ModelCoverageStep {...props} workloadClient={workloadClient} />
  );

  const ModeledDataStepWrapper = (props: WizardStepProps) => (
    <ModeledDataSetupStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
      phase="create"
    />
  );

  const ModelSummaryStepWrapper = (props: WizardStepProps) => (
    <ModelSummaryStep
      {...props}
      workloadClient={workloadClient}
      workspaceId={workspaceId}
    />
  );

  const steps: WizardStep[] = [
    {
      id: "overview",
      title: "Overview",
      description: "Create a semantic view of your device data",
      component: ModelOverviewStep,
      validate: () => true,
      numbered: false,
    },
    {
      id: "upload-model",
      title: "Device Model",
      description: "Upload a model and review field coverage",
      component: DeviceModelStepWrapper,
      validate: (context) =>
        !!context.modelUploaded &&
        [
          ...(context.modelTelemetries || []),
          ...(context.modelProperties || []),
        ].some((field: { included?: boolean }) => field.included),
    },
    {
      id: "raw-data",
      title: "IoT Data Source",
      description: "Select and scope the device data",
      component: RawDataStepWrapper,
      validate: (context) => !!context.modelRawDataSelected,
    },
    {
      id: "model-coverage",
      title: "Model coverage",
      description: "Compare observed fields with the device model",
      component: CoverageStepWrapper,
      validate: (context) => !!context.modelCoverageVerified,
    },
    {
      id: "create-entities",
      title: "Modeled data",
      description: "Create typed and state-enriched device data",
      component: ModeledDataStepWrapper,
      validate: (context) => !!context.modeledTablesCreated,
      isNavigationDisabled: (context) => !!context.modeledDataCreating,
    },
    {
      id: "summary",
      title: "Ready to Use",
      description: "Review the modeled data and next steps",
      component: ModelSummaryStepWrapper,
      numbered: false,
      validate: (context) => !!context.modeledTablesCreated,
    },
  ];

  return (
    <WizardControl
      title="Model Your Device Data"
      steps={steps}
      onComplete={onWizardComplete}
      canFinish={(_stepId, context) => !!context.modeledTablesCreated}
      showNavigation={true}
      allowStepNavigation={true}
      navigationLabels={{
        next: "Next",
        previous: "Previous",
        complete: "Finish and Close",
        cancel: "Cancel",
      }}
      onFirstStepBack={onBack}
      persistKey={`iot-wizard-v4-model-${itemId}`}
    />
  );
}
