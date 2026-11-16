// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import { Button, Checkbox, Dropdown, Input, Label, MessageBar, MessageBarBody, Option, Radio, RadioGroup, Spinner, Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow, Text, Textarea, Tooltip } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { CheckmarkCircle24Filled, DismissCircle24Filled, Info16Regular } from "@fluentui/react-icons";
import { WizardStepProps } from "../../../components/Wizard";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import { canonicalJson, executeKqlRequest, isJsonObject, type AuthoringSource, type SampleStream, type StreamSample } from "../ModelAuthoring.ts";
import {
  applyRoleHints, clearAutomaticHints, effectiveScope, fieldValidation, matchRoleEvidence, inspectFieldEvidence, inspectScopedCoverage, inspectStreamCoverage,
  nativeBindingIssues, processingBindingSamples, setFieldRole, modelValidation,
  type HubModel, type ModelField, type ModelingInput, type ProcessingScope,
} from "../HubModeling.ts";
import { assembleStreamCoverage, beginStreamCheck, cancelStreamCheck, coverageConditionExample, currentStreamChecks, currentStreamResult, editStreamChecks, finishStreamCheck, modelCoverageReady, type StreamTicket } from "../StreamCoverage.ts";
import { buildModeledDataPlan, type ModeledDataPlan } from "../ModeledDataPlan.ts";
import { creationFingerprint, ModeledDataCreator, rememberCreationAttempt, type CreationAttempt, type CreationRequest, type CreationResult } from "../ModeledDataCreation.ts";
import { currentSourceAnalysis, type SourceReadiness } from "../ModelSourceReadiness.ts";
import { currentAssignmentHubSummary, currentAssignmentIntent, modelReferenceFilterAvailability } from "../HubAssignment.ts";
import { declaredMetadataDetails, displayMetadataValue } from "../FieldSemantics.ts";
import "../IoTSolutionItem.scss";
export { HubModelAssignmentStep } from "./HubModelAssignmentStep";

interface WorkflowProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI; workspaceId?: string; sourceRevision?: number; inline?: boolean;
  sourceState?: SourceReadiness; getSourceState?: () => SourceReadiness;
}
export function hubSource(context: Record<string, any>): AuthoringSource {
  return { queryServiceUri: context.modelQueryServiceUri || "", databaseName: context.modelDbName || "",
    telemetryTable: context.rawTelemetryTable || "", twinTable: context.rawPropertiesTable || "" };
}
export function hubScope(context: Record<string, any>, sourceState?: SourceReadiness): ProcessingScope {
  const model: HubModel | undefined = context.hubModel;
  const scope: ProcessingScope = context.hubScope || { mode: "unselected", operator: "and", telemetry: "", twin: "" };
  const additional = scope.additional || "unselected";
  return { ...scope, additional, telemetry: additional === "specified" ? scope.telemetry : "",
    twin: additional === "specified" ? scope.twin : "",
    mode: modelReferenceFilterAvailability(context, sourceState) === "available" ? scope.mode : "data",
    reference: model?.id && !modelValidation(model).length ? { id: model.id } : undefined };
}
export function hubInput(context: Record<string, any>, sourceState?: SourceReadiness): ModelingInput {
  return { source: hubSource(context), model: context.hubModel, fields: context.hubFields || [],
    scope: hubScope(context, sourceState), limit: 100 };
}
function useHeading(focus = true) {
  const ref = useRef<HTMLHeadingElement>(null);
  useEffect(() => { if (focus) ref.current?.focus(); }, [focus]);
  return ref;
}
function MissingModel({ onGoToStep, wizardContext }: WizardStepProps) {
  return <div className="iot-solution-step"><MessageBar intent="warning"><MessageBarBody>
    {modelValidation(wizardContext.hubModel).join(" ")}
  </MessageBarBody></MessageBar><Button onClick={() => onGoToStep("device-model")}>Return to Device Model</Button></div>;
}
function modelSchemaLabel(schema: ModelField["schema"]): string {
  if (typeof schema === "string") return schema;
  if (Array.isArray(schema)) return schema.map(value => typeof value === "string" ? value : "Structured value").join(" or ");
  if (schema && typeof schema === "object") {
    const type = schema.type || schema["@type"];
    if (typeof type === "string") return type;
    if (Array.isArray(type)) return type.filter((value): value is string => typeof value === "string").join(" or ") || "Structured value";
  }
  return "Unspecified";
}
function FieldQualifier({ name }: { name: string }) {
  const parts = name.match(/[^./_]*[./_]|[^./_]+$/g) || [name];
  return <span className="iot-solution-field-qualifier">({parts.map((part, index) =>
    <React.Fragment key={index}>{index > 0 && <wbr />}{part}</React.Fragment>)})</span>;
}
export function HubModelFieldsStep(props: WorkflowProps) {
  return !modelValidation(props.wizardContext.hubModel).length ? <HubModelFieldsContent {...props} /> : <MissingModel {...props} />;
}
function HubModelFieldsContent({ wizardContext, updateContext, resetStepsFrom, stepIndex, workspaceId = "", inline = false }: WorkflowProps) {
  const heading = useHeading(!inline);
  const Heading = inline ? "h3" : "h2";
  const GroupHeading = inline ? "h4" : "h3";
  const model: HubModel = wizardContext.hubModel;
  const fields: ModelField[] = wizardContext.hubFields || [];
  const [visible, setVisible] = useState(100);
  const analysis = currentSourceAnalysis(workspaceId, wizardContext);
  const evidence = analysis ? matchRoleEvidence(fields, analysis.roleEvidence) : [];
  const suggestionKey = analysis ? JSON.stringify([analysis.key, analysis.selectionRevision, analysis.analyzedAt, model.document]) : "";
  const suggested = analysis ? applyRoleHints(fields, evidence) : clearAutomaticHints(fields);
  const invalidate = () => {
    resetStepsFrom(stepIndex + 1);
    updateContext("hubCreationHistory", rememberCreationAttempt(wizardContext.hubCreationHistory || [], wizardContext.hubCreation));
    for (const key of ["hubCreation", "hubCoverage", "hubStreamChecks", "hubPreview"]) updateContext(key, undefined);
    updateContext("hubAllowWaiting", false);
  };
  const updateFields = (next: ModelField[]) => { updateContext("hubFields", next); invalidate(); };
  useEffect(() => {
    if (model.kind !== "wot") return;
    if (canonicalJson(fields) !== canonicalJson(suggested)) updateFields(suggested);
    if ((wizardContext.hubRoleEvidenceKey || "") !== suggestionKey) {
      updateContext("hubRoleEvidence", analysis ? evidence : undefined);
      updateContext("hubRoleEvidenceKey", suggestionKey || undefined);
    }
  }, [suggestionKey, fields, updateContext]);
  const suggestionStatus = !analysis ? "Validate the source tables in Select data source to get role suggestions."
    : !analysis.telemetry.payloads.length && !analysis.twin.payloads.length
      ? "No sample records were found. Choose the roles manually."
      : evidence.some(item => item.state === "telemetry" || item.state === "twin")
        ? ""
        : "The analyzed data gives no clear role suggestions. Choose the roles manually; your explicit choices are kept.";
  const errors = fieldValidation(fields, { includeRoleErrors: model.kind !== "wot" });
  const roleId = (id: string) => `field-role-${Array.from(id).map(character => character.codePointAt(0)!.toString(16)).join("-")}`;
  const roleLabel = <>Role{model.kind === "dtdl" && <Tooltip content="Roles come from your DTDL model" relationship="description">
    <Button appearance="subtle" size="small" icon={<Info16Regular />} aria-label="About DTDL roles" />
  </Tooltip>}</>;
  return <div className={inline ? "iot-solution-model-fields-inline" : "iot-solution-step iot-solution-model-preparation"}>
    <Heading ref={heading} id="model-fields-heading" tabIndex={-1}
      className={inline ? "iot-solution-outcome-heading" : "iot-solution-step-title"}
      style={{ marginBottom: "var(--spacingVerticalS)" }}>{inline ? "3. Select fields" : "Select fields"}</Heading>
    <Text className="iot-solution-step-description" block>Choose the fields to include as typed columns. Fields found in your recent device events have a suggested role. Select a role for any others. Fields you leave out stay in your raw tables.</Text>
    <ul className="iot-solution-outcome-list">
      <li><strong>Telemetry:</strong> Readings that change over time, like temperature</li>
      <li><strong>Twin property:</strong> Device settings or state, like firmware version</li>
    </ul>
    <section className="iot-solution-config-card" aria-labelledby="model-fields-group-heading">
      <div className="iot-solution-model-group-header" style={{ marginBottom: "var(--spacingVerticalS)" }}>
        <div>
          <GroupHeading id="model-fields-group-heading" className="iot-solution-outcome-heading">Fields</GroupHeading>
          <Text size={200} block role="status" style={{ color: "var(--colorNeutralForeground3)" }}>
            Select a role for every field to continue
          </Text>
          {model.kind === "wot" && suggestionStatus && <Text size={200} block role="status"
            style={{ color: "var(--colorNeutralForeground3)" }}>{suggestionStatus}</Text>}
        </div>
        <Button appearance="subtle" size="small" onClick={() => updateFields(fields.map(field => ({ ...field, included: !fields.every(item => item.included) })))}>
          {fields.every(field => field.included) ? "Deselect all" : "Select all"}
        </Button>
      </div>
      <div className="iot-solution-field-list iot-solution-model-field-list">
        <div className="iot-solution-model-field-columns">
          <Text size={200} weight="semibold" data-field-column="field">Field</Text>
          <Text size={200} weight="semibold" data-field-column="type">Type</Text>
          <Text size={200} weight="semibold" data-field-column="role">{roleLabel}</Text>
        </div>
        {fields.slice(0, visible).map(field => {
          const name = field.component ? `${field.component}.${field.path.join(".")}` : field.path.join(".");
          const needed = model.kind === "wot" && field.included && !field.role;
          const id = roleId(field.id);
          const schemaLabel = modelSchemaLabel(field.schema);
          const details = field.declaredMetadata ? declaredMetadataDetails(field.declaredMetadata)
            .filter(({ name }) => ["type", "schema", "unit"].includes(name))
            .sort((left, right) => Number(left.name === "unit") - Number(right.name === "unit"))
            : schemaLabel === "Unspecified" ? [] : [{ name: "type", value: schemaLabel }];
          return <div key={field.id} className="iot-solution-field-list-item iot-solution-model-field-card"
            data-field-id={field.id} data-selection-needed={needed || undefined}>
            <Checkbox className="iot-solution-model-field-include" aria-label={`Include ${field.title}`} checked={field.included}
              onChange={(_, data) => updateFields(fields.map(item => item.id === field.id ? { ...item, included: data.checked === true } : item))} />
            <Text size={200} weight="semibold" className="iot-solution-model-field-mobile-label" data-field-column="field">Field</Text>
            <div className="iot-solution-model-field-content">
              <div className="iot-solution-model-field-heading">
                <Text size={300}>{field.title}</Text>{field.title !== name && <> <FieldQualifier name={name} /></>}
              </div>
              <div className="iot-solution-model-field-details">
                {field.description && <Text className="iot-solution-model-field-description" block>{field.description}</Text>}
              </div>
            </div>
            <div className="iot-solution-model-field-affordances">
              {details.length > 0 && <Text size={200} weight="semibold" className="iot-solution-model-field-mobile-label" data-field-column="type">Type</Text>}
              {details.length > 0 && <div className="iot-solution-model-field-metadata">
                <Text size={200} block>{details.map((detail, index) =>
                  <React.Fragment key={detail.name}>{index > 0 && " \u00b7 "}{
                    detail.name === "schema" && (isJsonObject(detail.value) || Array.isArray(detail.value))
                      ? modelSchemaLabel(detail.value) : typeof detail.value === "string" ? detail.value
                        : displayMetadataValue(detail.value)}</React.Fragment>)}</Text>
              </div>}
              <div className="iot-solution-model-field-role">
                <Text size={200} weight="semibold" className="iot-solution-model-field-mobile-label" data-field-column="role">{roleLabel}</Text>
                {model.kind === "dtdl" ? <Text size={200}>{field.role === "telemetry" ? "Telemetry" : "Twin Property"}</Text>
                : <><Dropdown id={id} className={needed ? "iot-solution-selection-needed" : undefined}
                  aria-label={`Model ${field.title} as`} disabled={!field.included} aria-invalid={needed}
                  aria-describedby={needed ? `${id}-hint` : undefined} placeholder="Choose a Role"
                  value={field.role === "telemetry" ? "Telemetry" : field.role === "twin" ? "Twin Property" : ""}
                  selectedOptions={field.role ? [field.role] : []}
                  onOptionSelect={(_, data) => {
                    const role = data.optionValue;
                    if (role !== "telemetry" && role !== "twin") return;
                    updateFields(fields.map(item => item.id === field.id ? setFieldRole(item, role) : item));
                  }}>
                  <Option value="telemetry">Telemetry</Option><Option value="twin">Twin Property</Option>
                </Dropdown>{needed && <span id={`${id}-hint`} className="iot-solution-selection-cue">
                  Select a role<span className="iot-solution-model-visually-hidden">{` for ${field.title}. Choose Telemetry or Twin property.`}</span>
                </span>}</>}
              </div>
            </div>
          </div>;
        })}
      </div>
      {fields.length > visible && <Button appearance="subtle" onClick={() => setVisible(value => value + 100)}>Show 100 more fields</Button>}
    </section>
    {errors.length > 0 && <MessageBar intent="warning" layout="multiline"><MessageBarBody><ul>
      {errors.map((issue, index) => <li key={index}>{issue}</li>)}
    </ul></MessageBarBody></MessageBar>}
  </div>;
}

export function HubModelScopeStep(props: WorkflowProps) {
  return !modelValidation(props.wizardContext.hubModel).length ? <HubModelScopeContent {...props} /> : <MissingModel {...props} />;
}
function HubModelScopeContent({ wizardContext, updateContext, workloadClient, resetStepsFrom, stepIndex, sourceRevision = 0, sourceState, getSourceState }: WorkflowProps) {
  const heading = useHeading();
  const input = hubInput(wizardContext, sourceState), scope = input.scope;
  const storedScope: ProcessingScope = wizardContext.hubScope || { mode: "unselected", operator: "and", telemetry: "", twin: "" };
  const availability = modelReferenceFilterAvailability(wizardContext, sourceState);
  const skippedAssignment = availability === "unconfirmed" && currentAssignmentIntent(wizardContext, sourceState)?.mode === "skip" &&
    !currentAssignmentHubSummary(wizardContext, sourceState);
  const sourceReady = sourceState?.ready !== false;
  const sourceIsCurrent = () => {
    const current = getSourceState?.();
    return !current || current.ready && current.key === sourceState?.key && current.revision === sourceState?.revision;
  };
  const checks = currentStreamChecks(input, sourceRevision, wizardContext.hubStreamChecks);
  const checksRef = useRef(checks); checksRef.current = checks;
  const inputRef = useRef(input); inputRef.current = input;
  const sourceRef = useRef(sourceRevision); sourceRef.current = sourceRevision;
  const requests = useRef<Partial<Record<SampleStream, AbortController>>>({});
  const ownedTickets = useRef<Partial<Record<SampleStream, StreamTicket>>>({});
  const mounted = useRef(true);
  const streams = ["telemetry", "twin"] as const;
  const fieldName = (field: ModelField) => field.component ? `${field.component}.${field.path.join(".")}` : field.path.join(".");
  const expressionError = (stream: SampleStream): string => {
    try { if (scope.mode !== "unselected" && scope.additional !== "unselected") effectiveScope(scope, stream); return ""; }
    catch (failure: unknown) { return failure instanceof Error ? failure.message : "Review this filter."; }
  };
  const expressionErrors = { telemetry: expressionError("telemetry"), twin: expressionError("twin") };
  const fieldErrors = fieldValidation(input.fields);
  const checking = streams.some(stream => checks[stream].phase === "checking");
  const results = { telemetry: currentStreamResult(input, sourceRevision, checks, "telemetry"),
    twin: currentStreamResult(input, sourceRevision, checks, "twin") };
  const outstanding = streams.filter(stream => !results[stream]);
  const verificationFailures = new Map<string, SampleStream[]>();
  for (const stream of streams) {
    const check = checks[stream];
    if (check.phase === "error") {
      const message = check.error || "The data request failed.";
      verificationFailures.set(message, [...(verificationFailures.get(message) || []), stream]);
    }
  }
  const showCoverage = sourceReady && sourceIsCurrent() && !verificationFailures.size && streams.some(stream => !!results[stream] ||
    checks[stream].phase === "checking" && ownedTickets.current[stream]?.base === checks.base &&
      ownedTickets.current[stream]?.generation === checks[stream].generation);
  const publish = (next: typeof checks, currentInput = inputRef.current) => {
    checksRef.current = next;
    updateContext("hubStreamChecks", next);
    updateContext("hubCoverage", assembleStreamCoverage(currentInput, sourceRef.current, next));
  };
  useLayoutEffect(() => {
    if (checks !== wizardContext.hubStreamChecks) {
      for (const stream of ["telemetry", "twin"] as const) requests.current[stream]?.abort();
      publish(checks);
      updateContext("hubAllowWaiting", false);
      invalidate();
    }
  }, [checks]);
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      let next = checksRef.current;
      for (const stream of ["telemetry", "twin"] as const) {
        const ticket = ownedTickets.current[stream];
        if (!ticket) continue;
        requests.current[stream]?.abort();
        next = cancelStreamCheck(next, ticket);
      }
      ownedTickets.current = {};
      if (next !== checksRef.current) publish(next);
    };
  }, []);
  const invalidate = () => {
    resetStepsFrom(stepIndex + 1);
    updateContext("hubCreationHistory", rememberCreationAttempt(wizardContext.hubCreationHistory || [], wizardContext.hubCreation));
    updateContext("hubCreation", undefined); updateContext("hubCoverage", undefined); updateContext("hubPreview", undefined);
    updateContext("hubAllowWaiting", false);
  };
  const changeScope = (next: ProcessingScope) => {
    const effectiveNext = hubScope({ ...wizardContext, hubScope: next }, sourceState);
    const ledger = editStreamChecks(input, effectiveNext, sourceRevision, checksRef.current);
    for (const stream of ["telemetry", "twin"] as const) {
      if (ledger.revision !== checksRef.current.revision || ledger[stream].generation !== checksRef.current[stream].generation) requests.current[stream]?.abort();
    }
    checksRef.current = ledger;
    updateContext("hubStreamChecks", ledger); updateContext("hubScope", next); invalidate();
  };
  const inspect = async (stream: SampleStream,
    authorize = () => acquireTokenWithConsent(workloadClient, "https://kusto.kusto.windows.net/.default")) => {
    if (!sourceReady || !sourceIsCurrent() || scope.mode === "unselected" || scope.additional === "unselected" || fieldErrors.length ||
        expressionErrors[stream] || checksRef.current[stream].phase === "checking") return;
    requests.current[stream]?.abort();
    const controller = new AbortController(); requests.current[stream] = controller;
    const started = beginStreamCheck(input, sourceRevision, stream, checksRef.current);
    ownedTickets.current[stream] = started.ticket;
    invalidate(); publish(started.checks);
    try {
      const token = await authorize();
      if (controller.signal.aborted || !mounted.current) return;
      if (!sourceIsCurrent()) {
        publish(cancelStreamCheck(checksRef.current, started.ticket)); return;
      }
      const result = await inspectStreamCoverage(input, token.token, stream, controller.signal);
      if (controller.signal.aborted || !mounted.current) return;
      if (!sourceIsCurrent()) {
        publish(cancelStreamCheck(checksRef.current, started.ticket)); return;
      }
      const live = currentStreamChecks(inputRef.current, sourceRef.current, checksRef.current);
      publish(finishStreamCheck(inputRef.current, live, started.ticket, { result }));
    } catch (failure: unknown) {
      if (!controller.signal.aborted && mounted.current) {
        if (!sourceIsCurrent()) { publish(cancelStreamCheck(checksRef.current, started.ticket)); return; }
        const live = currentStreamChecks(inputRef.current, sourceRef.current, checksRef.current);
        publish(finishStreamCheck(inputRef.current, live, started.ticket,
          { error: failure instanceof Error ? failure.message : "The data request failed." }));
      }
    } finally {
      if (ownedTickets.current[stream] === started.ticket) delete ownedTickets.current[stream];
    }
  };
  const verify = () => {
    if (streams.some(stream => checksRef.current[stream].phase === "checking")) return;
    let authorization: ReturnType<typeof acquireTokenWithConsent> | undefined;
    const authorize = () => authorization ||= acquireTokenWithConsent(workloadClient, "https://kusto.kusto.windows.net/.default");
    const needed = streams.filter(stream => !currentStreamResult(inputRef.current, sourceRef.current, checksRef.current, stream));
    for (const stream of needed.length ? needed : streams) void inspect(stream, authorize);
  };
  const snapshot = assembleStreamCoverage(input, sourceRevision, checks);
  const ready = modelCoverageReady(input, sourceRevision, checks);
  const bindings = snapshot && processingBindingSamples(snapshot);
  const issues = bindings ? nativeBindingIssues(bindings.telemetry, bindings.twin) : [];
  const nextAction = !sourceReady ? sourceState?.reason || "Validate the current source tables before checking coverage."
    : fieldErrors.length ? "Confirm your selected fields and their roles in Device Model before checking coverage."
    : scope.mode === "unselected" ? "Choose a model ID filter before checking coverage."
    : scope.additional === "unselected" ? "Choose whether to add conditions before checking model coverage."
    : checking ? "Checking coverage. Results are reported separately for each stream."
    : outstanding.length ? "Choose Check coverage to check streams that don't have current results. Resolve any failed checks before continuing."
      : ready ? "Model coverage check complete."
        : "Review the results and resolve the source or binding issues before continuing.";
  const coverageRows = streams.map(stream => {
    const check = checks[stream], result = results[stream], sample = result?.sample;
    const empty: StreamSample = { rawRows: 0, observations: 0, skippedWithoutReportedProperties: 0, payloads: [] };
    const evidence = sample ? inspectFieldEvidence(input.fields, stream === "telemetry" ? sample : empty, stream === "twin" ? sample : empty) : [];
    const fields = input.fields.filter(field => field.included && field.role === stream);
    const missing = fields.filter(field => evidence.some(item => item.fieldId === field.id && item[stream] === 0));
    return { stream, check, result, sample, fields, missing };
  });
  return <div className="iot-solution-step iot-solution-model-preparation iot-solution-model-coverage">
    <h2 ref={heading} tabIndex={-1} className="iot-solution-step-title">Filter events and check model coverage</h2>
    <Text className="iot-solution-step-description">Your raw tables can include events from other kinds of devices. Choose which events to include in your modeled data, then check how well your model covers them.</Text>
    <section className="iot-solution-outcome-section" aria-labelledby="model-reference-filter-heading">
      <h3 id="model-reference-filter-heading" className="iot-solution-outcome-heading">1. Filter by model ID</h3>
      {availability === "available" ? <>
        <Text block>Events from devices assigned to this model include its model ID. Choose whether to use the ID to filter events.</Text>
        <RadioGroup className="iot-solution-coverage-choices" aria-labelledby="model-reference-filter-heading"
          value={scope.mode === "reference" ? scope.operator : scope.mode === "data" ? "data" : ""}
          onChange={(_, data) => {
            if (data.value === "data") changeScope({ ...storedScope, mode: "data" });
            else if (data.value === "and" || data.value === "or") changeScope({ ...storedScope, mode: "reference", operator: data.value });
          }}>
          <div className="iot-solution-coverage-choice" data-selected={scope.mode === "data"}>
            <Radio value="data" aria-labelledby="reference-none-label" aria-describedby="reference-none-description" label={<span>
              <span id="reference-none-label" className="iot-solution-coverage-choice-title">Don't filter by model ID</span>
              <span id="reference-none-description" className="iot-solution-coverage-choice-description">Use only the additional conditions to choose events</span>
            </span>} />
          </div>
          <div className="iot-solution-coverage-choice" data-selected={scope.mode === "reference" && scope.operator === "and"}>
            <Radio value="and" aria-labelledby="reference-and-label" aria-describedby="reference-and-description" label={<span>
              <span id="reference-and-label" className="iot-solution-coverage-choice-title">Require the model ID</span>
              <span id="reference-and-description" className="iot-solution-coverage-choice-description">Include only events with this model's ID ({input.model.id}) that also meet any additional conditions</span>
            </span>} />
          </div>
          <div className="iot-solution-coverage-choice" data-selected={scope.mode === "reference" && scope.operator === "or"}>
            <Radio value="or" aria-labelledby="reference-or-label" aria-describedby="reference-or-description" label={<span>
              <span id="reference-or-label" className="iot-solution-coverage-choice-title">Model ID or additional conditions</span>
              <span id="reference-or-description" className="iot-solution-coverage-choice-description">Include events with this model's ID, plus any events that meet the additional conditions</span>
            </span>} />
          </div>
        </RadioGroup>
      </> : <MessageBar intent={availability === "unlinked" ? "warning" : "info"} layout="multiline"
        bottomReflowSpacer={null} className="iot-solution-assignment-notice">
        <MessageBarBody>{availability === "dtdl"
          ? "Model ID filtering isn't available for DTDL models. To enable model ID filtering, go back and upload a WoT Thing Model."
          : availability === "unlinked"
            ? "None of the IoT Hubs in your dataset are linked to an ADR namespace. Model reference filtering is not available for these IoT Hubs."
            : skippedAssignment ? "Model ID filtering isn't available because model assignment was skipped. To enable model ID filtering, go back and assign the model to your devices."
            : "Model reference filtering is unavailable until your IoT Hub links are checked and confirmed. IoT Hub links are checked after choosing Assign model to devices on the Assign model page."}
        </MessageBarBody>
      </MessageBar>}
    </section>
    <section className="iot-solution-outcome-section iot-solution-coverage-conditions" aria-labelledby="additional-conditions-heading">
    <h3 id="additional-conditions-heading" className="iot-solution-outcome-heading">2. Add conditions (optional)</h3>
    <Text id="coverage-conditions-description" size={200} block>Filter events further with a KQL condition for each raw table</Text>
    <RadioGroup className="iot-solution-coverage-choices" aria-labelledby="additional-conditions-heading"
      value={scope.additional === "none" || scope.additional === "specified" ? scope.additional : ""}
      onChange={(_, data) => {
        if (data.value === "none" || data.value === "specified") changeScope({ ...storedScope, additional: data.value });
      }}>
      <div className="iot-solution-coverage-choice" data-selected={scope.additional === "none"}>
        <Radio value="none" label="No additional conditions" />
      </div>
      <div className="iot-solution-coverage-choice" data-selected={scope.additional === "specified"}>
        <Radio value="specified" label="Add conditions" />
      </div>
    </RadioGroup>
    {scope.additional === "specified" && <>
    <Text id="coverage-predicate-description" block>Enter a KQL condition, like the part after where in a query. Only events that match are included.</Text>
    <div className="iot-solution-coverage-editors">
    {streams.map(stream => {
      const example = coverageConditionExample(input.fields, stream, true);
      const table = stream === "telemetry" ? input.source.telemetryTable : input.source.twinTable;
      return <section key={stream} className="iot-solution-stream-filter" aria-label={stream === "telemetry" ? "Telemetry conditions" : "Reported properties conditions"}>
        <div className="iot-solution-field"><Label htmlFor={`hub-${stream}-scope`}>{stream === "telemetry" ? "Telemetry table" : "Properties table"} <span className="iot-solution-coverage-optional">(optional)</span></Label>
          <Textarea id={`hub-${stream}-scope`} value={scope[stream]} rows={5} resize="vertical"
            placeholder={`Condition for ${table}\nExample: ${example}`} aria-describedby="coverage-predicate-description"
            aria-invalid={!!expressionErrors[stream]}
            onChange={(_, data) => changeScope({ ...storedScope, [stream]: data.value })} /></div>
        {expressionErrors[stream] && <Text role="alert">{expressionErrors[stream]}</Text>}
      </section>;
    })}
    </div>
    </>}
    </section>
    <section className="iot-solution-outcome-section iot-solution-coverage-results" aria-labelledby="coverage-results-heading">
    <h3 id="coverage-results-heading" className="iot-solution-outcome-heading">3. Check model coverage</h3>
    <Text id="coverage-sample-note" className="iot-solution-step-description" block>Check up to 100 recent matching events from each table to see which of your selected fields appear</Text>
    <div className="iot-solution-model-actions">
      <Button appearance="primary" disabled={!sourceReady || scope.mode === "unselected" || scope.additional === "unselected" || !!fieldErrors.length ||
        checking || streams.some(stream => !!expressionErrors[stream])} onClick={verify}
        icon={checking ? <Spinner size="tiny" /> : undefined}>Check coverage</Button>
    </div>
    {sourceReady && sourceIsCurrent() && verificationFailures.size > 0 &&
      <MessageBar intent="error" role="alert" layout="multiline" className="iot-solution-coverage-error"><MessageBarBody>
        {[...verificationFailures].map(([message, failedStreams]) => <Text key={message} block>
          <strong>{failedStreams.map(stream => stream === "telemetry" ? "Telemetry" : "Reported Properties").join(" and ")}: </strong>{message}
        </Text>)}
      </MessageBarBody></MessageBar>}
    {showCoverage && sourceReady && sourceIsCurrent() && ready && scope.mode === "reference" &&
      !!snapshot?.sourceProbe?.telemetry.payloads.length && !!snapshot.sourceProbe.twin.payloads.length &&
      !nativeBindingIssues(snapshot.sourceProbe.telemetry, snapshot.sourceProbe.twin).length &&
      coverageRows.some(row => row.sample?.payloads.length === 0 && !!row.result?.probe.payloads.length) &&
      <MessageBar intent="info" role="note" layout="multiline" bottomReflowSpacer={null}
        className="iot-solution-coverage-assignment-hint"><MessageBarBody>
        New model assignments can take some time to show up in device events
      </MessageBarBody></MessageBar>}
    {showCoverage && <div className="iot-solution-coverage-streams" aria-describedby="coverage-sample-note">
      {coverageRows.map(({ stream, check, result, sample, fields, missing }) => {
        const label = stream === "telemetry" ? "Telemetry table" : "Properties table";
        const status = check.phase === "error" ? "Couldn't verify."
          : check.phase === "checking" ? "Verifying..."
          : !sample ? check.edited ? "Verify again." : "Not verified."
          : !sample.payloads.length ? "No matching events"
          : !fields.length ? "No fields selected"
          : missing.length ? `${missing.length} field${missing.length === 1 ? "" : "s"} not found` : "All fields found";
        const allObserved = !!sample?.payloads.length && fields.length > 0 && missing.length === 0;
        return <section key={stream} className="iot-solution-coverage-stream" data-coverage-stream={stream}
          aria-labelledby={`coverage-${stream}-heading`} aria-busy={check.phase === "checking"}>
          <div className="iot-solution-coverage-stream-heading">
            <h4 id={`coverage-${stream}-heading`}>{label}{" "}
              <Text className="iot-solution-coverage-source">({stream === "telemetry" ? input.source.telemetryTable : input.source.twinTable})</Text>
            </h4>
          </div>
          <dl className="iot-solution-coverage-details">
            <div><dt>Matching events checked</dt>
            <dd data-sample-count aria-label={sample ? `Matching events checked: ${sample.payloads.length} (limit ${input.limit}).` : "Not verified"}>
              <span className="iot-solution-coverage-number">{sample ? sample.payloads.length : "\u2014"}</span>
            </dd></div>
            <div><dt>Fields found</dt><dd data-observed-fields aria-label={sample ? `Selected model fields found: ${fields.length - missing.length} of ${fields.length} in this sample.` : "Not verified"}>
              {sample ? fields.length ? <><span className="iot-solution-coverage-number">{fields.length - missing.length}</span><span className="iot-solution-coverage-denominator"> of {fields.length}</span></> : "None selected" : "\u2014"}
            </dd></div>
            <div><dt>Status</dt><dd><Text role="status" className="iot-solution-coverage-result-status"
              data-success={allObserved} data-attention={!!sample && (!sample.payloads.length || missing.length > 0)}>{status}</Text></dd></div>
          </dl>
          {sample && (!sample.payloads.length || missing.length > 0) && <>
          {sample.payloads.length === 0 && <Text size={200} block>{result!.probe.rawRows === 0
            ? "The source sample was empty. Confirm that data is arriving in the selected raw table."
            : result!.probe.payloads.length === 0
              ? "No usable payloads were observed in the source sample. Confirm the selected raw table contains supported updates."
              : "No events match your filters. Check your conditions, or wait for devices to send matching data."}</Text>}
          {sample.payloads.length > 0 && missing.length > 0 && <div className="iot-solution-coverage-missing">
            <Table size="extra-small" className="iot-solution-coverage-missing-fields">
              <caption>Fields not found in these events ({missing.length})</caption>
              <TableHeader className="iot-solution-model-visually-hidden"><TableRow>
                <TableHeaderCell scope="col">Field</TableHeaderCell>
              </TableRow></TableHeader>
              <TableBody>{missing.map(field => {
              const technicalName = fieldName(field), title = field.title || technicalName;
              return <TableRow key={field.id} data-missing-field={field.id}>
                <TableCell>
                  <Text className="iot-solution-coverage-missing-title">{title}</Text>
                  {title !== technicalName && <>{" ("}<Text className="iot-solution-coverage-missing-name">{technicalName}</Text>{")"}</>}
                </TableCell>
              </TableRow>;
            })}</TableBody></Table>
            <Text size={200} block data-missing-fields-guidance>You can still continue. These fields stay empty until devices send them.</Text>
          </div>}
          </>}
        </section>;
      })}
    </div>}
    {issues.length > 0 && <MessageBar intent="warning" layout="multiline"><MessageBarBody>{issues.join(" ")}</MessageBarBody></MessageBar>}
    <Text role="status" className="iot-solution-model-visually-hidden">{nextAction}</Text>
    </section>
  </div>;
}
export function HubModelReviewStep(props: WorkflowProps) {
  return !modelValidation(props.wizardContext.hubModel).length ? <HubModelReviewContent {...props} /> : <MissingModel {...props} />;
}
function HubModelReviewContent({ wizardContext, updateContext, workloadClient, resetStepsFrom, stepIndex, sourceRevision = 0, sourceState }: WorkflowProps) {
  const heading = useHeading(), input = hubInput(wizardContext, sourceState);
  const coverage = assembleStreamCoverage(input, sourceRevision, wizardContext.hubStreamChecks);
  const eligible = modelCoverageReady(input, sourceRevision, wizardContext.hubStreamChecks);
  const [names, setNames] = useState<ModeledDataPlan["names"]>(() => {
    if (wizardContext.hubOutputNames) return wizardContext.hubOutputNames;
    const stem = (input.model.title || "Model").replace(/[^a-zA-Z0-9_]/g, "_").slice(0, 60), suffix = crypto.randomUUID().slice(0, 6);
    return { modeled: `ModeledData_${stem}_${suffix}`, twinChanges: `PropertiesNormalized_${stem}_${suffix}`, twinState: `PropertiesLKV_${stem}_${suffix}` };
  });
  useEffect(() => { updateContext("hubOutputNames", names); }, [names, updateContext]);
  const [creator] = useState(() => new ModeledDataCreator());
  const [running, setRunning] = useState(false);
  const busy = running || !!wizardContext.hubCreationBusy;
  const busyRef = useRef(busy); busyRef.current = busy;
  const [operationError, setOperationError] = useState("");
  const submitting = useRef(false), mounted = useRef(true);
  const operationController = useRef<AbortController>();
  useEffect(() => {
    mounted.current = true;
    let active = true;
    workloadClient.navigation.onBeforeNavigateAway(async () => ({ canLeave: !active || (!submitting.current && !busyRef.current) }));
    return () => { active = false; mounted.current = false; operationController.current?.abort(); };
  }, [workloadClient]);
  const request: CreationRequest = { input, coverage, allowWaiting: true, names };
  const fingerprint = creationFingerprint(request), prior: CreationResult | undefined = wizardContext.hubCreation;
  const current = prior?.fingerprint === fingerprint ? prior : undefined, completed = eligible && current?.phase === "completed";
  const history: CreationAttempt[] = rememberCreationAttempt(wizardContext.hubCreationHistory || [], prior);
  const historyRef = useRef(history); historyRef.current = history;
  let plan: ModeledDataPlan | undefined, planError = "";
  try {
    if (!eligible) throw new Error("Check telemetry and reported properties before creating modeled data.");
    plan = buildModeledDataPlan(input, coverage, true, names);
  } catch (failure: unknown) { planError = failure instanceof Error ? failure.message : "Review your setup."; }
  const groups = new Map<string, { label: string; commands: ModeledDataPlan["commands"]; metadata: boolean }>();
  for (const command of plan?.commands || []) {
    const metadata = command.kind === "tableDoc" || command.kind === "viewDoc" || command.kind === "columnDocs";
    const key = `${metadata ? "docstrings" : command.kind}:${command.target}`;
    const group = groups.get(key) || { label: metadata ? `Set docstring on ${command.target}` : command.label, commands: [], metadata };
    group.commands.push(command); groups.set(key, group);
  }
  const operationRows = [...groups.values()].sort((a, b) => Number(a.metadata) - Number(b.metadata)).map(group => {
    const steps = group.commands.map(command => current?.steps.find(step => step.label === command.label));
    const needsFinalReadback = group.commands.some(command => command.kind === "columnDocs");
    const finalCheck = current?.steps.find(step => step.label === "Verify all created definitions and native semantic documentation");
    const readback = needsFinalReadback || finalCheck?.status === "failed" &&
      group.commands.some(command => finalCheck.detail?.includes(command.target)) ? finalCheck : undefined;
    const failed = steps.find(step => step?.status === "failed") || (readback?.status === "failed" ? readback : undefined);
    const passed = steps.every(step => step?.status === "passed") && (!needsFinalReadback || readback?.status === "passed");
    const started = steps.some(Boolean);
    const command = group.commands[0];
    const successLabel = group.metadata ? `Added description to ${command.target}`
      : command.kind === "policy" ? `Added update policy to ${command.target}`
        : command.kind === "table" || command.kind === "view" ? `Created ${command.target}` : group.label;
    const status = failed ? "failed" : passed ? "passed" : started ? "running" : "pending";
    return { label: group.label, status, displayLabel: status === "passed" ? successLabel : group.label, detail: failed?.detail };
  });
  const visibleHistory = history.filter(attempt => attempt.operationId !== current?.operationId || attempt.phase === "failed");
  const run = async () => {
    if (submitting.current || busy || completed || !plan || !eligible) return;
    submitting.current = true; busyRef.current = true; setRunning(true); setOperationError("");
    const controller = new AbortController(); operationController.current = controller;
    updateContext("hubCreationBusy", true); resetStepsFrom(stepIndex + 1);
    try {
      const token = await acquireTokenWithConsent(workloadClient, "https://kusto.kusto.windows.net/.default");
      if (!mounted.current || controller.signal.aborted) return;
      await creator.run(request, {
        request: (csl, kind) => {
          if (!mounted.current || controller.signal.aborted) throw new Error("The creation view closed. Any dispatched operation needs review before retrying.");
          return executeKqlRequest(input.source, token.token, csl, kind, controller.signal);
        },
        inspect: fresh => inspectScopedCoverage(fresh, token.token, "processing", undefined, controller.signal),
      }, true, result => {
        updateContext("hubCreation", result);
        if (result.coverage) updateContext("hubCoverage", result.coverage);
        historyRef.current = rememberCreationAttempt(historyRef.current, result); updateContext("hubCreationHistory", historyRef.current);
      });
    } catch (failure: unknown) {
      if (mounted.current) setOperationError(failure instanceof Error ? failure.message : "Couldn't create modeled data.");
    } finally {
      submitting.current = false; busyRef.current = false; updateContext("hubCreationBusy", false);
      if (mounted.current) setRunning(false);
    }
  };
  const editNames = (next: ModeledDataPlan["names"]) => {
    updateContext("hubCreationHistory", history); updateContext("hubCreation", undefined); setNames(next);
  };
  const telemetryCount = input.fields.filter(field => field.included && field.role === "telemetry").length;
  const twinCount = input.fields.filter(field => field.included && field.role === "twin").length;
  return <div className="iot-solution-step">
    <h2 ref={heading} tabIndex={-1} className="iot-solution-step-title">Create modeled data</h2>
    <Text className="iot-solution-step-description">Review the resources this step creates in your KQL database. Together, they turn raw JSON into typed columns and add each device's latest state to its telemetry.</Text>
    <div className="iot-solution-form">
      <div className="iot-solution-validation">
        <div className="iot-solution-validation-row"><Text className="iot-solution-validation-label">Telemetry table</Text>
          <Text className="iot-solution-validation-value">{input.source.telemetryTable}</Text></div>
        <div className="iot-solution-validation-row"><Text className="iot-solution-validation-label">Properties table</Text>
          <Text className="iot-solution-validation-value">{input.source.twinTable}</Text></div>
        <div className="iot-solution-validation-row"><Text className="iot-solution-validation-label">Selected fields</Text>
          <Text className="iot-solution-validation-value">{telemetryCount} telemetry, {twinCount} twin {twinCount === 1 ? "property" : "properties"}</Text></div>
      </div>
      <Text id="creation-resources-heading" weight="semibold" size={400} block style={{ marginTop: "12px" }}>Resources to create</Text>
      <MessageBar intent="info"><MessageBarBody>Property updates flow into the normalized table, the latest values are kept in a materialized view, and the modeled data table combines them with telemetry</MessageBarBody></MessageBar>
      <Text id="output-names-help" className="iot-solution-resource-name-note">You can rename these resources before you create them</Text>
      {([
        ["modeled", "modeled-output-name", "Modeled data table",
          `Your typed data: ${telemetryCount} telemetry column${telemetryCount === 1 ? "" : "s"} and ${twinCount} property column${twinCount === 1 ? "" : "s"}, plus device ID and event time, for each event`],
        ["twinChanges", "twin-changes-name", "Normalized properties table", "Stores each property update as its own row: device, property, value, and time. Use it for property history"],
        ["twinState", "twin-state-name", "Latest property values (materialized view)", "Keeps each device's latest reported value for every property"],
      ] as const).map(([key, id, label, help]) =>
        <div key={key} className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor={id}>{label}</Label>
          <Input id={id} required aria-describedby={`output-names-help ${id}-help`} value={names[key]} disabled={busy || completed}
            onChange={(_, data) => editNames({ ...names, [key]: data.value })} />
          <Text id={`${id}-help`} className="iot-solution-field-hint">{help}</Text>
        </div>)}
      {planError && <MessageBar intent="error" role="alert"><MessageBarBody>{planError}</MessageBarBody></MessageBar>}
      <Button appearance="primary" disabled={busy || completed || !plan} onClick={run} style={{ marginTop: "8px" }}>
        {busy ? "Creating modeled data..." : "Create modeled data"}</Button>
    </div>
    {operationError && <MessageBar intent="error" role="alert"><MessageBarBody>{operationError}</MessageBarBody></MessageBar>}
    {(current?.phase === "blocked" || current?.phase === "failed") && current.error &&
      <MessageBar intent="error" role="alert"><MessageBarBody>{current.error}</MessageBarBody></MessageBar>}
    {current && current.steps.length > 0 && <div className="iot-solution-creation-steps" role="list" aria-label="Setup progress"
      aria-live="polite" aria-busy={busy}>
      {operationRows.map((step, index) => <div key={index} className="iot-solution-creation-step" role="listitem" data-status={step.status}>
        {step.status === "running" ? <Spinner size="tiny" /> : step.status === "passed"
          ? <CheckmarkCircle24Filled aria-hidden="true" primaryFill="var(--colorPaletteGreenForeground1)" />
          : step.status === "failed" ? <DismissCircle24Filled aria-hidden="true" primaryFill="var(--colorPaletteRedForeground1)" />
          : <span className="iot-solution-step-dot" aria-hidden="true" />}
        <div className="iot-solution-creation-step-content">
          <Text className={`iot-solution-creation-step-label ${step.status === "passed" ? "success" : step.status === "failed" ? "error" : step.status}`}>
            {step.displayLabel}</Text>
          {step.detail && <Text className="iot-solution-field-error">{step.detail}</Text>}
        </div>
      </div>)}
    </div>}
    {completed && <MessageBar intent="success" role="status" id="modeled-data-success" layout="multiline" bottomReflowSpacer={null}><MessageBarBody>
      <Text className="iot-solution-success-text">Modeled data created. Select <strong>Continue</strong> to see what was created and what to do next.</Text>
    </MessageBarBody></MessageBar>}
    {visibleHistory.length > 0 && <details className="iot-solution-model-details" open={current?.phase === "failed"}>
      <summary>Attempt history ({visibleHistory.length})</summary>
      {visibleHistory.map((attempt, index) => <section key={attempt.operationId} className="iot-solution-history-attempt" aria-label={`Attempt ${index + 1}: ${attempt.names.modeled}`}>
        <h4>Attempt {index + 1}: {attempt.names.modeled}</h4>
        <Text size={200} block>Database: {attempt.databaseName}. Output names: {attempt.names.modeled}, {attempt.names.twinChanges}, {attempt.names.twinState}.</Text>
        {attempt.phase === "failed" ? <MessageBar intent="warning" layout="multiline"><MessageBarBody>
          <strong>Setup didn't finish. Some changes may have been made. Review the attempt details before trying again.</strong>{" "}{attempt.error}
          <div>Nothing was automatically rolled back.</div>
          {attempt.uncertainCommand && <div>The outcome of {attempt.uncertainCommand} is unconfirmed. Check it in {attempt.databaseName} before trying again.</div>}
          {attempt.objects.length === 0
            ? <div>No created objects were confirmed; this does not prove that nothing was created.</div>
            : <div>Review the confirmed objects and use unused output names for another attempt. Existing names will not be reused.</div>}
        </MessageBarBody></MessageBar> : <Text size={200} block>
          {attempt.operationId === current?.operationId && completed ? "Setup complete" : "This earlier completed attempt does not verify the current setup."}
        </Text>}
        {attempt.objects.length > 0 && <><Text weight="semibold">Confirmed objects</Text>
          <ul>{attempt.objects.map(object => <li key={object.name}>{object.name} ({object.kind})</li>)}</ul></>}
      </section>)}
    </details>}
  </div>;
}
