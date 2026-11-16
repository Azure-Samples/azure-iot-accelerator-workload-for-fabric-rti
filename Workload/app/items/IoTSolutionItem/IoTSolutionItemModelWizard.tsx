// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useCallback, useMemo, useRef, useState } from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardControl, WizardStepProps } from "../../components/Wizard";
import { ItemWithDefinition } from "../../controller/ItemCRUDController";
import { IoTSolutionItemDefinition } from "./IoTSolutionItemDefinition";
import { HubModelSourceStep } from "./steps/HubModelSourceStep";
import { ModelPreparationStep } from "./steps/ModelPreparationStep";
import { ModelOverviewStep } from "./steps/ModelOverviewStep";
import { HubModelConfiguredStep } from "./steps/HubModelConfiguredStep";
import { HubModelAssignmentStep, HubModelScopeStep, HubModelReviewStep, hubInput } from "./steps/HubModelWorkflowSteps";
import { fieldValidation, modelValidation } from "./HubModeling.ts";
import { creationFingerprint, creationOutputsConfirmed, rememberCreationAttempt, type CreationResult } from "./ModeledDataCreation.ts";
import { currentSourceAnalysis, modelSourceKey, SOURCE_PENDING, type SourceReadiness, type SourceSelectionStatus } from "./ModelSourceReadiness.ts";
import { modelCoverageReady } from "./StreamCoverage.ts";
import { assignmentIntentOwner, currentAssignmentIntent, currentAssignmentListsLoading } from "./HubAssignment.ts";
import "./IoTSolutionItem.scss";

interface IoTSolutionItemModelWizardProps {
  workloadClient: WorkloadClientAPI;
  item?: ItemWithDefinition<IoTSolutionItemDefinition>;
  onBack?: () => void;
  initialContext?: Record<string, unknown>;
}

export function IoTSolutionItemModelWizard({
  workloadClient,
  item,
  onBack,
  initialContext,
}: IoTSolutionItemModelWizardProps) {
  const workspaceId = item?.workspaceId || "";
  const source = useRef<SourceReadiness>({ key: "", ready: false, reason: SOURCE_PENDING, revision: 0 });
  const activeProps = useRef<WizardStepProps>();
  const [, renderSource] = useState(0);
  const publishSource = useCallback((next: SourceSelectionStatus) => {
    const previous = source.current;
    if (previous.key === next.key && previous.ready === next.ready && previous.reason === next.reason) return;
    source.current = { ...next, revision: previous.revision + 1 };
    const props = activeProps.current;
    if (props && (previous.key !== next.key || previous.ready && !next.ready)) {
      const context = props.wizardContext;
      props.resetStepsFrom(Math.min(props.stepIndex + 1, 3));
      props.updateContext("hubAssignmentIntent", undefined);
      props.updateContext("hubAssignmentHubSummary", undefined);
      props.updateContext("hubCreationHistory", rememberCreationAttempt(context.hubCreationHistory || [], context.hubCreation));
      for (const key of ["hubCreation", "hubCoverage", "hubStreamChecks", "hubPreview"]) props.updateContext(key, undefined);
      props.updateContext("hubAllowWaiting", false);
    }
    renderSource(value => value + 1);
  }, []);
  const views = useMemo(() => {
    const bind = (props: WizardStepProps) => {
      // React's development stack probes can invoke a step factory without render props.
      if (props && Object.prototype.hasOwnProperty.call(props, "wizardContext")) activeProps.current = props;
    };
    const currentSource = () => source.current.key === modelSourceKey(workspaceId, activeProps.current?.wizardContext || {}) &&
      !!currentSourceAnalysis(workspaceId, activeProps.current?.wizardContext || {})
      ? source.current : { ...source.current, ready: false, reason: SOURCE_PENDING };
    return {
        Overview: (props: WizardStepProps) => <ModelOverviewStep {...props} workloadClient={workloadClient} />,
      Source: (props: WizardStepProps) => {
        bind(props);
        return <HubModelSourceStep {...props} workloadClient={workloadClient} workspaceId={workspaceId} onSourceStateChange={publishSource} />;
      },
      Model: (props: WizardStepProps) => {
        bind(props);
        return <ModelPreparationStep {...props} workloadClient={workloadClient} workspaceId={workspaceId}
          sourceState={currentSource()} getSourceState={currentSource} />;
      },
      Assignment: (props: WizardStepProps) => {
        bind(props);
        return <HubModelAssignmentStep {...props} workloadClient={workloadClient}
          sourceState={currentSource()} getSourceState={currentSource} />;
      },
      Scope: (props: WizardStepProps) => { bind(props); return <HubModelScopeStep {...props} workloadClient={workloadClient}
        sourceRevision={source.current.revision} sourceState={currentSource()} getSourceState={currentSource} />; },
      Review: (props: WizardStepProps) => { bind(props); return <HubModelReviewStep {...props} workloadClient={workloadClient}
        sourceRevision={source.current.revision} sourceState={currentSource()} getSourceState={currentSource} />; },
    };
  }, [workloadClient, workspaceId, publishSource]);
  const Configured = (props: WizardStepProps) => <HubModelConfiguredStep {...props} workloadClient={workloadClient} workspaceId={workspaceId}
    sourceState={source.current} sourceRevision={source.current.revision} />;
  const assignmentReady = (context: Record<string, any>) => source.current.key === modelSourceKey(workspaceId, context) &&
    !!currentSourceAnalysis(workspaceId, context) && !fieldValidation(context.hubFields || []).length &&
    !!assignmentIntentOwner(context, source.current) && !context.hubAssignmentBusy &&
    !currentAssignmentListsLoading(context, source.current) &&
    (context.hubModel?.kind === "dtdl" || !!currentAssignmentIntent(context, source.current));
  const verifiedCreation = (context: Record<string, any>) => {
    const result = context.hubCreation as CreationResult | undefined;
    if (!source.current.ready || !currentSourceAnalysis(workspaceId, context) || !assignmentReady(context) ||
        modelValidation(context.hubModel).length || !context.hubOutputNames || result?.phase !== "completed" ||
        !modelCoverageReady(hubInput(context, source.current), source.current.revision, context.hubStreamChecks)) return false;
    return creationOutputsConfirmed(result, { input: hubInput(context, source.current), names: context.hubOutputNames }) &&
      result.fingerprint === creationFingerprint({
      input: hubInput(context, source.current), names: context.hubOutputNames,
      coverage: context.hubCoverage, allowWaiting: true,
    });
  };
  return (
    <div>
      <WizardControl title="Model your device data" className="iot-solution-hub-model-wizard" steps={[
        { id: "overview", title: "Get started", description: "See what modeling unlocks",
          component: views.Overview, numbered: false, validate: () => true },
        { id: "raw-data", title: "Select data source", description: "Choose your raw IoT tables",
          component: views.Source, validate: context => !!context.modelRawDataSelected && !!currentSourceAnalysis(workspaceId, context) &&
            source.current.ready && source.current.key === modelSourceKey(workspaceId, context) },
        { id: "device-model", title: "Add device model", description: "Upload a model and choose fields",
          component: views.Model, validate: context => !!context.modelRawDataSelected && source.current.ready &&
            source.current.key === modelSourceKey(workspaceId, context) && !!currentSourceAnalysis(workspaceId, context) &&
            !modelValidation(context.hubModel).length && fieldValidation(context.hubFields || []).length === 0 },
        { id: "assign-references", title: "Assign model", description: "Optional: link devices to the model",
          component: views.Assignment, isNavigationDisabled: context => !!context.hubAssignmentBusy, validate: assignmentReady },
        { id: "scope-coverage", title: "Check model coverage", description: "Pick devices and confirm coverage",
          component: views.Scope, validate: context => !!context.hubModel && assignmentReady(context) &&
            modelCoverageReady(hubInput(context, source.current), source.current.revision, context.hubStreamChecks) },
        { id: "review-plan", title: "Create modeled data", description: "Build typed telemetry with device state",
          component: views.Review, isNavigationDisabled: context => !!context.hubCreationBusy, validate: verifiedCreation },
        { id: "configured", title: "Review results", description: "See what was created and next steps",
          component: Configured, numbered: false, validate: verifiedCreation },
      ]}
        onComplete={() => onBack?.()} canFinish={(stepId, context) => stepId === "configured" && verifiedCreation(context)}
        initialContext={initialContext}
        initialStepId={initialContext?.modelRawDataSelected ? "raw-data" : "overview"}
        onFirstStepBack={onBack}
        navigationLabels={{ next: "Continue", complete: "Finish and Close", previous: "Previous" }}
      />
    </div>
  );
}
