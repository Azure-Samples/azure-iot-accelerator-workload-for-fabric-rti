// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardControl, WizardStep, WizardStepProps } from "../../components/Wizard";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { DashboardOverviewStep } from "./steps/DashboardOverviewStep";
import { DashboardDataStep } from "./steps/DashboardDataStep";
import { DashboardGenerateStep } from "./steps/DashboardGenerateStep";
import { DashboardSummaryStep } from "./steps/DashboardSummaryStep";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemDashboardWizardProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onWizardComplete: (context: Record<string, any>) => void;
  onBack?: () => void;
}

/**
 * Dashboard generation wizard.
 * Step 1: Select a data source and inspect the table fields.
 * Step 2: Generate a Real-Time Dashboard from the selected fields.
 */
export function IoTSolutionItemDashboardWizard({
  workloadClient,
  item,
  onWizardComplete,
  onBack,
}: IoTSolutionItemDashboardWizardProps) {
  const workspaceId = item?.workspaceId || "";
  const itemId = item?.id || "default";

  // Wrap DashboardDataStep to inject extra props
  const DashboardDataStepWrapper = (props: WizardStepProps) => (
    <DashboardDataStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  const DashboardGenerateStepWrapper = (props: WizardStepProps) => (
    <DashboardGenerateStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  const DashboardSummaryStepWrapper = (props: WizardStepProps) => (
    <DashboardSummaryStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} />
  );

  const steps: WizardStep[] = [
    {
      id: "overview",
      title: "Overview",
      description: "Visualize live device behavior and trends",
      component: DashboardOverviewStep,
      validate: () => true,
      numbered: false,
    },
    {
      id: "data-source",
      title: "Data Source",
      description: "Select raw data tables and inspect fields",
      component: DashboardDataStepWrapper,
      validate: (context) => !!context.dashInspected,
    },
    {
      id: "generate-dashboard",
      title: "Generate Dashboard",
      description: "Create the Real-Time Dashboard",
      component: DashboardGenerateStepWrapper,
      validate: (context) => !!context.dashDashboardCreated,
      isNavigationDisabled: (context) => !!context.dashDashboardCreating,
    },
    {
      id: "summary",
      title: "Ready to Use",
      description: "Use and extend your dashboard",
      component: DashboardSummaryStepWrapper,
      validate: (context) => !!context.dashDashboardCreated,
      numbered: false,
    },
  ];

  return (
    <WizardControl
      title="Generate Real-Time Dashboards"
      steps={steps}
      onComplete={onWizardComplete}
      canFinish={(_stepId, context) => !!context.dashDashboardCreated}
      showNavigation={true}
      allowStepNavigation={true}
      navigationLabels={{
        next: "Next",
        previous: "Previous",
        complete: "Finish and Close",
        cancel: "Cancel",
      }}
      onFirstStepBack={onBack}
      persistKey={`iot-wizard-v4-dashboard-${itemId}`}
    />
  );
}
