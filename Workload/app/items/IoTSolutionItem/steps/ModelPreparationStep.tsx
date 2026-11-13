// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import {
  Button, MessageBar, MessageBarBody, Radio, RadioGroup, Spinner, Text,
} from "@fluentui/react-components";
import { ArrowUpload24Regular, Checkmark16Regular, DocumentCheckmark24Regular } from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { AuthoringSample, inferRootPreparationSchema, isJsonObject, type JsonValue } from "../ModelAuthoring.ts";
import { resolveThingModel, TM_NAMESPACE } from "../ThingModelParser.ts";
import { parseDtdlModel, UnsupportedDtdlModelFileError } from "../DtdlModelParser.ts";
import { applyRoleHints, clearAutomaticHints, fromDtdlModel, fromThingModel, matchRoleEvidence, modelValidation,
  type HubModel, type ModelField } from "../HubModeling.ts";
import { rememberCreationAttempt } from "../ModeledDataCreation.ts";
import { currentSourceAnalysis, SOURCE_PENDING, sourceActionCurrent, type SourceReadiness } from "../ModelSourceReadiness.ts";
import { HubModelFieldsStep } from "./HubModelWorkflowSteps";
import "../IoTSolutionItem.scss";

interface ModelPreparationStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId?: string;
  sourceState?: SourceReadiness;
  getSourceState?: () => SourceReadiness;
}

const UNSUPPORTED_MODEL_RECOVERY = "This file isn't a supported device model. Upload a WoT Thing Model, a DTDL model, or an IoT Central export.";
const COPILOT_PROMPT_PREFIX = "Help me create and enrich a valid WoT 1.1 Thing Model with an absolute model ID and a valid JSON-LD context for the following JSON Schema:";

export function ModelPreparationStep({
  wizardContext, updateContext, resetStepsFrom, stepIndex, workloadClient, workspaceId = "",
  sourceState = { ready: false, reason: SOURCE_PENDING, key: "", revision: 0 }, getSourceState,
  ...wizardProps
}: ModelPreparationStepProps) {
  const retainedModelErrors = wizardContext.hubModel ? modelValidation(wizardContext.hubModel) : [];
  const [format, setFormat] = useState<string>(() => {
    const selected = wizardContext.hubModel?.kind || wizardContext.hubSelectedFormat;
    return selected === "wot" || selected === "dtdl" ? selected : "wot";
  });
  const [sample, setSample] = useState<AuthoringSample | null>(null);
  const [reading, setReading] = useState(false);
  const [sampleError, setSampleError] = useState("");
  const [emptySample, setEmptySample] = useState(false);
  const [preparationOpen, setPreparationOpen] = useState(false);
  const [sampleOwner, setSampleOwner] = useState("");
  const [copyState, setCopyState] = useState<"idle" | "copying" | "copied" | "failed">("idle");
  const [fileName, setFileName] = useState<string>(wizardContext.preparationFileName || "");
  const [modelError, setModelError] = useState(retainedModelErrors.join(" "));
  const [modelErrorDetail, setModelErrorDetail] = useState("");
  const [readingFile, setReadingFile] = useState(false);
  const [hubModel, setHubModel] = useState<HubModel | undefined>(() => retainedModelErrors.length ? undefined : wizardContext.hubModel);
  const [rootDocument, setRootDocument] = useState<JsonValue | undefined>(wizardContext.hubModel?.document || wizardContext.hubRootDocument);
  const [dragging, setDragging] = useState(false);
  const request = useRef<AbortController | null>(null);
  const fileRevision = useRef(0);
  const heading = useRef<HTMLHeadingElement>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const preparationTrigger = useRef<HTMLButtonElement>(null);
  const preparationRegion = useRef<HTMLDivElement>(null);
  const promptElement = useRef<HTMLPreElement>(null);
  const preparationWasOpen = useRef(false);
  const copyGeneration = useRef(0);
  const copyResetTimer = useRef<number>();
  const invalidateCopyFeedback = () => {
    copyGeneration.current++;
    if (copyResetTimer.current !== undefined) window.clearTimeout(copyResetTimer.current);
    copyResetTimer.current = undefined;
  };
  const analysis = currentSourceAnalysis(workspaceId, wizardContext);
  const sourceKey = JSON.stringify([sourceState.key, sourceState.revision, sourceState.ready,
    analysis?.key, analysis?.selectionRevision, analysis?.analyzedAt]);
  const liveSourceKey = useRef(sourceKey); liveSourceKey.current = sourceKey;
  const prompt = sourceState.ready && analysis && sample && sampleOwner === sourceKey
    ? `${COPILOT_PROMPT_PREFIX}\n\n${JSON.stringify(sample.schema, null, 2)}` : "";
  const currentPrerequisite = useRef(sourceState); currentPrerequisite.current = sourceState;
  const currentSourceState = () => getSourceState?.() || currentPrerequisite.current;
  const retained = useRef({ hubModel, rootDocument, fileName, fields: wizardContext.hubFields || [] });
  const filePending = useRef(false);
  if (hubModel && !filePending.current) retained.current = { hubModel, rootDocument, fileName, fields: wizardContext.hubFields || [] };
  useLayoutEffect(() => {
    request.current?.abort();
    invalidateCopyFeedback();
    setCopyState("idle"); setSampleOwner(""); setPreparationOpen(false);
    setSample(null); setSampleError(""); setReading(false); setEmptySample(false); setDragging(false);
    if (filePending.current) {
      fileRevision.current++;
      filePending.current = false; setReadingFile(false);
      const previous = retained.current;
      setHubModel(previous.hubModel); setRootDocument(previous.rootDocument); setFileName(previous.fileName);
      updateContext("hubModel", previous.hubModel); updateContext("hubFields", clearAutomaticHints(previous.fields));
      updateContext("hubRootDocument", previous.rootDocument); updateContext("preparationFileName", previous.fileName);
      setModelError(""); setModelErrorDetail("");
    }
    return () => { request.current?.abort(); invalidateCopyFeedback(); };
  }, [sourceKey]);
  useLayoutEffect(() => {
    if (preparationOpen && !preparationWasOpen.current) preparationRegion.current?.focus();
    else if (!preparationOpen && preparationWasOpen.current) preparationTrigger.current?.focus();
    if (preparationOpen && prompt && document.activeElement === preparationRegion.current) promptElement.current?.focus();
    preparationWasOpen.current = preparationOpen;
  }, [preparationOpen, prompt]);

  useEffect(() => () => {
    fileRevision.current++;
    if (filePending.current) {
      const previous = retained.current;
      updateContext("hubModel", previous.hubModel); updateContext("hubFields", clearAutomaticHints(previous.fields));
      updateContext("hubRootDocument", previous.rootDocument); updateContext("preparationFileName", previous.fileName);
    }
  }, []);
  useEffect(() => { heading.current?.focus(); }, []);
  useEffect(() => {
    if (wizardContext.hubModel && modelValidation(wizardContext.hubModel).length) {
      invalidateCoverage();
      setModelError(modelValidation(wizardContext.hubModel).join(" "));
      setHubModel(undefined);
      updateContext("hubModel", undefined);
      retained.current = { ...retained.current, hubModel: undefined };
    }
  }, [wizardContext.hubModel]);

  const preparePrompt = async () => {
    const prerequisite = currentSourceState();
    if (!prerequisite.ready) return;
    request.current?.abort();
    const controller = new AbortController();
    request.current = controller;
    const owner = sourceKey;
    invalidateCopyFeedback(); setCopyState("idle"); setSampleOwner("");
    setReading(true);
    setSample(null);
    setSampleError("");
    setEmptySample(false);
    try {
      if (!analysis) throw new Error("Choose Validate source tables in Select data source to analyze device data first.");
      // Yield a paint before inferring the bounded cached sample.
      await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
      if (controller.signal.aborted || liveSourceKey.current !== owner || !sourceActionCurrent(prerequisite, currentSourceState())) return;
      const value: AuthoringSample = { limit: analysis.limit, telemetry: analysis.telemetry, twin: analysis.twin,
        schema: inferRootPreparationSchema(analysis.telemetry, analysis.twin) };
      if (!controller.signal.aborted && liveSourceKey.current === owner && sourceActionCurrent(prerequisite, currentSourceState())) {
        setSample(value); setSampleOwner(owner);
      }
    } catch (error: unknown) {
      if (!controller.signal.aborted && sourceActionCurrent(prerequisite, currentSourceState())) {
        const message = error instanceof Error ? error.message : "We couldn't prepare the prompt. Try again.";
        if (message.startsWith("No object payload samples are available.")) setEmptySample(true);
        else setSampleError(message);
      }
    } finally {
      if (!controller.signal.aborted && sourceActionCurrent(prerequisite, currentSourceState())) setReading(false);
    }
  };

  const invalidateCoverage = () => {
    resetStepsFrom(stepIndex + 1);
    updateContext("hubCreationHistory", rememberCreationAttempt(wizardContext.hubCreationHistory || [], wizardContext.hubCreation));
    updateContext("hubCreation", undefined);
    updateContext("hubCoverage", undefined);
    updateContext("hubStreamChecks", undefined);
    updateContext("hubPreview", undefined);
    updateContext("hubAllowWaiting", false);
  };
  const beginModelChange = () => {
    invalidateCoverage();
    updateContext("hubModelRevision", (wizardContext.hubModelRevision || 0) + 1);
    updateContext("hubAssignmentIntent", undefined);
    updateContext("hubAssignmentHubSummary", undefined);
    updateContext("hubAssignmentChoices", undefined); updateContext("hubAssignmentInventory", undefined);
    updateContext("hubRoleEvidence", undefined); updateContext("hubRoleEvidenceKey", undefined);
    updateContext("hubHintCounts", undefined);
    updateContext("hubHintBindings", undefined);
    updateContext("hubScope", { ...(wizardContext.hubScope || { telemetry: "", twin: "", operator: "and" }), mode: "unselected", additional: "unselected", reference: undefined });
    updateContext("hubModel", undefined);
  };
  const validateDocument = (document: JsonValue, name: string) => {
    if (format !== "wot" && format !== "dtdl") throw new Error("Choose a model format first");
    const types = (Array.isArray(document) ? document : [document])
      .flatMap(member => isJsonObject(member) ? [member["@type"]].flat() : []);
    const isThingModel = types.includes("tm:ThingModel") || types.includes(`${TM_NAMESPACE}ThingModel`);
    const isInterface = (value: JsonValue) => isJsonObject(value) && [value["@type"]].flat().includes("Interface");
    const isDtdl = (Array.isArray(document) ? document : [document]).some(isInterface) ||
      (isJsonObject(document) && isJsonObject(document.capabilityModel) && isInterface(document.capabilityModel));
    if (isJsonObject(document) && "$schema" in document && !isThingModel && !isDtdl) {
      throw new Error("This is a JSON Schema, not a device model. Upload a model file instead.");
    }
    if ((format === "wot" && isDtdl) || (format === "dtdl" && isThingModel)) {
      throw new Error("This file doesn't match the format you chose. Change the format or upload a different file.");
    }
    let next: HubModel;
    if (format === "wot") {
      if (!isJsonObject(document) && !Array.isArray(document)) throw new Error(UNSUPPORTED_MODEL_RECOVERY);
      const result = resolveThingModel(document, [], name);
      if (result.blockers.length) throw new Error(result.blockers.join(" "));
      next = fromThingModel(result);
      updateContext("preparationThingModel", result);
    } else {
      next = fromDtdlModel(parseDtdlModel(document), document);
      updateContext("preparationThingModel", null);
    }
    if (!next.fields.length) throw new Error(`No telemetry or property fields were found in ${name}. Upload a model that describes your device data.`);
    const previous: ModelField[] = wizardContext.hubFields || [];
    let fields = next.fields.map(field => {
      const existing = previous.find(item => item.id === field.id);
      return existing
        ? { ...field, included: existing.included, column: existing.column,
          ...(field.roleSource !== "model" && existing.roleSource === "manual" ? { role: existing.role, roleSource: existing.roleSource } : {}) }
        : field;
    });
    if (next.kind === "wot" && analysis) {
      const evidence = matchRoleEvidence(fields, analysis.roleEvidence);
      fields = applyRoleHints(fields, evidence);
      updateContext("hubRoleEvidence", evidence);
      updateContext("hubRoleEvidenceKey", JSON.stringify([analysis.key, analysis.selectionRevision, analysis.analyzedAt, next.document]));
    }
    setHubModel(next);
    updateContext("hubModel", next);
    updateContext("hubFields", fields);
    updateContext("hubSelectedFormat", format);
  };
  const resetModel = (name = "") => {
    const revision = ++fileRevision.current;
    beginModelChange();
    setModelError("");
    setModelErrorDetail("");
    setFileName(name);
    setRootDocument(undefined);
    updateContext("hubRootDocument", undefined);
    updateContext("preparationThingModel", null);
    updateContext("preparationFileName", name);
    setHubModel(undefined);
    setReadingFile(false);
    return revision;
  };
  const upload = async (file: File) => {
    const prerequisite = currentSourceState();
    if (!prerequisite.ready || filePending.current) return;
    const revision = resetModel(file.name);
    filePending.current = true;
    setReadingFile(true);
    let readFailureDetail = "";
    try {
      if (file.size > 2 * 1024 * 1024) throw new Error("Upload a file that's 2 MB or smaller");
      let text: string;
      try { text = await file.text(); } catch (error: unknown) {
        readFailureDetail = error instanceof Error ? error.message : "";
        throw new Error("This file couldn't be read. Upload a different model file.");
      }
      if (revision !== fileRevision.current || !sourceActionCurrent(prerequisite, currentSourceState())) return;
      let document: JsonValue;
      try { document = JSON.parse(text); } catch {
        throw new Error(`${file.name} isn't valid JSON. Fix the file, then upload it again.`);
      }
      setRootDocument(document);
      updateContext("hubRootDocument", document);
      validateDocument(document, file.name);
    } catch (error: unknown) {
      if (revision === fileRevision.current && sourceActionCurrent(prerequisite, currentSourceState())) {
        retained.current = { hubModel: undefined, rootDocument: undefined, fileName: file.name, fields: [] };
        const message = error instanceof Error ? error.message : "This file couldn't be read. Upload a different model file.";
        if (error instanceof UnsupportedDtdlModelFileError) {
          setModelError(UNSUPPORTED_MODEL_RECOVERY);
        } else if (message.startsWith("This model imports definitions with tm:ref") ||
            message.startsWith("This model extends another model with tm:extends")) {
          setModelError(message);
        } else if (/Required model (file|fragment) missing|was not supplied|not supplied locally/i.test(message)) {
          setModelError(`${file.name} refers to definitions that aren't included. Include them in this file, then choose it again.`);
          setModelErrorDetail(message);
        } else if (/not a WoT Thing Model/i.test(message)) {
          setModelError(UNSUPPORTED_MODEL_RECOVERY);
          setModelErrorDetail(message);
        } else setModelError(message);
        if (readFailureDetail) setModelErrorDetail(readFailureDetail);
      }
    } finally {
      if (revision === fileRevision.current && sourceActionCurrent(prerequisite, currentSourceState())) {
        filePending.current = false; setReadingFile(false);
      }
    }
  };
  const modelFileAction = hubModel || fileName ? "Replace Model" : "Choose File";
  const ignoredParts = hubModel?.kind === "wot" ? hubModel.ignored || [] : [];
  const chooseModelFile = () => {
    if (!filePending.current && currentSourceState().ready) fileInput.current?.click();
  };
  const collapsePreparation = () => {
    request.current?.abort(); setReading(false);
    invalidateCopyFeedback(); setCopyState("idle");
    setPreparationOpen(false);
  };
  const copyPrompt = async () => {
    const prerequisite = currentSourceState(), owner = sourceKey;
    if (!prompt || !preparationOpen || format !== "wot" || copyState === "copying" ||
        liveSourceKey.current !== owner || !sourceActionCurrent(sourceState, prerequisite)) return;
    invalidateCopyFeedback();
    const ticket = copyGeneration.current;
    setCopyState("copying");
    const current = () => ticket === copyGeneration.current && liveSourceKey.current === owner &&
      sourceActionCurrent(prerequisite, currentSourceState());
    try {
      await navigator.clipboard.writeText(prompt);
      if (current()) {
        setCopyState("copied");
        copyResetTimer.current = window.setTimeout(() => {
          copyResetTimer.current = undefined;
          if (current()) setCopyState("idle");
        }, 2000);
      }
    } catch {
      if (current()) setCopyState("failed");
    }
  };

  return (
    <section className="iot-solution-step iot-solution-model-preparation-section" aria-labelledby="model-acquisition-heading">
      <div>
        <h2 ref={heading} id="model-acquisition-heading" tabIndex={-1} className="iot-solution-step-title">Add a device model</h2>
        <Text className="iot-solution-step-description">Upload a model that describes your device data. You'll use it to choose which fields become typed columns.</Text>
      </div>
      {!sourceState.ready && <Text id="model-source-prerequisite" block role="status">{sourceState.reason}</Text>}
      <section id="model-upload-heading" tabIndex={-1} className="iot-solution-model-upload" aria-label="Model format and file">
        <h3 id="model-format-heading" className="iot-solution-outcome-heading">1. Choose a model format</h3>
        <RadioGroup className="iot-solution-coverage-choices iot-solution-model-format-choices"
          aria-labelledby="model-format-heading" layout="horizontal" value={format} disabled={!sourceState.ready}
          onChange={(_, data) => {
            if (!currentSourceState().ready || data.value === format) return;
            collapsePreparation();
            filePending.current = false;
            retained.current = { hubModel: undefined, rootDocument: undefined, fileName: "", fields: [] };
            updateContext("hubFields", []);
            resetModel();
            setFormat(data.value);
            updateContext("hubSelectedFormat", data.value);
          }}>
          <div className="iot-solution-coverage-choice" data-selected={format === "wot"}>
            <Radio value="wot" aria-labelledby="model-format-wot-title" aria-describedby="model-format-wot-description" label={<span>
              <span id="model-format-wot-title" className="iot-solution-coverage-choice-title">WoT Thing Model (recommended)</span>
              <span id="model-format-wot-description" className="iot-solution-coverage-choice-description">Lets you assign the model to devices.</span>
            </span>} />
          </div>
          <div className="iot-solution-coverage-choice" data-selected={format === "dtdl"}>
            <Radio value="dtdl" aria-labelledby="model-format-dtdl-title" aria-describedby="model-format-dtdl-description" label={<span>
              <span id="model-format-dtdl-title" className="iot-solution-coverage-choice-title">Digital Twin Definition Language (DTDL)</span>
              <span id="model-format-dtdl-description" className="iot-solution-coverage-choice-description">Use an existing DTDL model or IoT Central template.</span>
            </span>} />
          </div>
        </RadioGroup>
      {format === "wot" && <section className="iot-solution-config-card" aria-labelledby="model-preparation-heading">
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "8px" }}>
          <h3 id="model-preparation-heading" className="iot-solution-model-guide-heading">
            <Text weight="semibold" size={400}>Need help creating a Thing Model?</Text>
          </h3>
        </div>
        <Text className="iot-solution-step-description" block style={{ marginBottom: "12px" }}>
          Copilot can draft a Thing Model from your device data. Get a prompt to copy into Copilot in Fabric.
        </Text>
        {!preparationOpen && <Button ref={preparationTrigger} appearance="primary" aria-expanded={false}
          disabled={!sourceState.ready} aria-describedby={!sourceState.ready ? "model-source-prerequisite" : undefined}
          aria-controls="model-preparation-content" onClick={() => {
            if (currentSourceState().ready) { setPreparationOpen(true); void preparePrompt(); }
          }}>Get Copilot prompt</Button>}
        <div ref={preparationRegion} id="model-preparation-content" className="iot-solution-model-guide-content"
          role="region" tabIndex={-1} aria-labelledby="model-preparation-heading" hidden={!preparationOpen}>
          {reading && <Spinner size="small" label="Preparing prompt..." />}
          {sampleError && <MessageBar intent="error" role="alert"><MessageBarBody>{sampleError}</MessageBarBody></MessageBar>}
          {emptySample && <Text className="iot-solution-model-prompt-empty" role="status" block>No sample records were found during source validation. Choose Validate again in Select data source after data arrives.</Text>}
          {prompt && <>
            {sample && (!sample.telemetry.payloads.length || !sample.twin.payloads.length) && <Text size={200} block>
              Only {sample.telemetry.payloads.length ? "telemetry" : "reported-property"} sample records were available.
              {" "}This schema reflects the available stream; an empty stream does not prove that a field is unsupported.
            </Text>}
            <ol id="model-preparation-instructions">
              <li>Copy the prompt. It lists the fields found in your device data</li>
              <li>Paste it into Copilot in Fabric</li>
              <li>Review the Thing Model Copilot drafts, then save it as a .json file</li>
              <li>Upload the file in <strong>Upload your model</strong></li>
            </ol>
            <pre ref={promptElement} id="model-preparation-prompt" className="iot-solution-model-code" tabIndex={0}
              aria-label="Copilot prompt">{prompt}</pre>
          </>}
          <div className="iot-solution-model-actions">
            {prompt && <Button appearance="primary" disabled={!sourceState.ready} disabledFocusable={copyState === "copying"}
              aria-live="polite" aria-atomic="true" aria-busy={copyState === "copying"}
              icon={copyState === "copied" ? <Checkmark16Regular aria-hidden="true" /> : undefined}
              onClick={() => { void copyPrompt(); }}>{copyState === "copying" ? "Copying..." : copyState === "copied" ? "Copied" : "Copy prompt"}</Button>}
            <Button appearance="secondary" onClick={collapsePreparation}>Hide prompt</Button>
          </div>
          <Text className="iot-solution-model-prompt-disclaimer" size={200} block>Copilot can make mistakes. Check the model before you upload it.</Text>
          {copyState === "failed" && <Text role="alert" block>Copy failed. Select and copy the prompt shown above.</Text>}
        </div>
      </section>}
        {format && <>
        <Text id="model-supported-heading" className="iot-solution-field-label" weight="semibold" block>What's supported</Text>
        <ul className="iot-solution-outcome-list" aria-label={format === "wot" ? "Supported WoT models" : "Supported DTDL models"}>
          {format === "wot" ? <>
            <li><strong>Supported model files:</strong> a self-contained WoT 1.1 Thing Model object or an array of related Thing Models. Each model needs its own supplied absolute ID.</li>
            <li><strong>Supported capabilities:</strong> properties and event data using primitive, enum, and complex schemas.</li>
            <li><strong>Not included:</strong> actions, submodels, and external definitions.</li>
          </> : <>
            <li><strong>Supported model files:</strong> a single DTDL interface, an array of related interfaces, or an IoT Central device template export, including inherited capabilities and components.</li>
            <li><strong>Supported capabilities:</strong> root, inherited, and component telemetry and device-reported properties using primitive, enum, and complex schemas.</li>
            <li><strong>Not included:</strong> commands and cloud-only properties, because they are not present in device event data.</li>
          </>}
        </ul>
        <h3 id="model-file-upload-heading" className="iot-solution-outcome-heading" tabIndex={-1}>2. Upload your model</h3>
        <input ref={fileInput} id="thing-model-upload" hidden type="file"
          aria-label="Device model file" accept=".json,.jsonld,application/json,application/tm+json" disabled={readingFile || !sourceState.ready}
          onChange={event => {
            const file = event.currentTarget.files?.[0];
            event.currentTarget.value = "";
            if (file && currentSourceState().ready) void upload(file);
          }} />
        <div className={`iot-solution-model-upload-area iot-solution-drop-zone${hubModel ? " iot-solution-model-upload-area--loaded iot-solution-drop-zone--loaded" : ""}${dragging ? " iot-solution-drop-zone--active" : ""}`}
          role="button" aria-label={modelFileAction} tabIndex={readingFile || !sourceState.ready ? -1 : 0}
          aria-busy={readingFile}
          aria-disabled={readingFile || !sourceState.ready}
          aria-describedby={modelError ? "model-upload-error" : !sourceState.ready ? "model-source-prerequisite"
            : hubModel ? "model-upload-file-name model-upload-reference model-upload-replace-help" : undefined}
          onClick={chooseModelFile}
          onKeyDown={event => {
            if (event.key === "Enter" || event.key === " ") {
              event.preventDefault();
              if (!event.repeat) chooseModelFile();
            }
          }}
          onDragOver={event => { event.preventDefault(); if (!readingFile && currentSourceState().ready) setDragging(true); }}
          onDragLeave={event => { if (!(event.relatedTarget instanceof Node) || !event.currentTarget.contains(event.relatedTarget)) setDragging(false); }}
          onDrop={event => {
            event.preventDefault(); setDragging(false);
            if (readingFile || !currentSourceState().ready) return;
            const files = event.dataTransfer.files;
            if (files.length === 1) void upload(files[0]);
            else if (files.length > 1) {
              retained.current = { hubModel: undefined, rootDocument: undefined, fileName: "", fields: [] };
              resetModel(); setModelError("Upload one model file that includes all its definitions");
            }
          }}>
          <div className="iot-solution-model-upload-identity" aria-live="polite">
            {hubModel ? <>
              <DocumentCheckmark24Regular aria-hidden="true" style={{ color: "var(--colorPaletteGreenForeground1)", fontSize: "32px" }} />
              <Text id="model-upload-file-name" weight="semibold" block>{fileName}</Text>
              <Text size={200} block style={{ color: "var(--colorNeutralForeground3)", minWidth: 0, overflowWrap: "anywhere" }}>
                Model ID: <span id="model-upload-reference">{hubModel.id || "Not provided in the uploaded model."}</span>
              </Text>
              <Text id="model-upload-replace-help" size={200} style={{ color: "var(--colorNeutralForeground3)" }} block>
                Click or drop a new file to replace
              </Text>
            </> : readingFile ? <Spinner size="small" label={`Checking ${fileName}...`} /> : <>
              <ArrowUpload24Regular aria-hidden="true" />
              <Text weight="semibold" block>Drop a model file here</Text>
              <Text size={200} block>Self-contained JSON or JSON-LD, up to 2 MB.</Text>
              <Text size={200} block>{fileName ? "Click or drop a new file to replace" : "Click to choose a file, or drop it here."}</Text>
            </>}
          </div>
        </div>
        {ignoredParts.length > 0 && <MessageBar id="model-upload-warnings" className="iot-solution-model-upload-warning" intent="warning" layout="multiline"
          bottomReflowSpacer={null}><MessageBarBody>
          <Text block>Some parts of your model aren't supported and won't be included:</Text>
          <ul style={{ margin: "var(--spacingVerticalXS) 0 0", paddingInlineStart: "var(--spacingHorizontalL)" }}>
          {(["null type", "oneOf", "submodel"] as const).map(reason => {
            const parts = ignoredParts.filter(part => part.reason === reason);
            return parts.length ? <li key={reason}>
              {reason === "null type" ? "Fields with no type" : reason === "oneOf" ? "Fields that use oneOf" : "Submodels"}: {parts.map(part =>
                part.kind === "Submodel" ? part.name : `${part.kind === "Property" ? "property" : "event"} ${part.name}`).join(", ")}
            </li> : null;
          })}
          </ul>
        </MessageBarBody></MessageBar>}
        {modelError && <MessageBar id="model-upload-error" className="iot-solution-model-upload-error" intent="error" role="alert"
          aria-labelledby={fileName ? "model-upload-error-file" : undefined} aria-describedby="model-upload-error-message"><MessageBarBody>
          {fileName && <Text id="model-upload-error-file" weight="semibold" block>{fileName} couldn't be loaded. Try uploading it again</Text>}
          <div id="model-upload-error-message">
            {modelError}
            {modelErrorDetail && <details className="iot-solution-model-details"><summary>Validation details</summary>{modelErrorDetail}</details>}
          </div>
        </MessageBarBody></MessageBar>}
        {hubModel && <HubModelFieldsStep {...wizardProps} wizardContext={wizardContext} updateContext={updateContext}
          resetStepsFrom={resetStepsFrom} stepIndex={stepIndex} workloadClient={workloadClient} workspaceId={workspaceId} inline />}
        </>}
      </section>
    </section>
  );
}
