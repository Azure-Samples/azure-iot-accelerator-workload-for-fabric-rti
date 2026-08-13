import React, { useState, useCallback } from "react";
import {
  Button,
  Checkbox,
  Input,
  Label,
  Text,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  ArrowUpload24Regular,
  DocumentCheckmark24Regular,
} from "@fluentui/react-icons";
import { WizardStepProps } from "../../../components/Wizard";
import { parseDtdlModel, DtdlCapability, DtdlParseResult } from "../DtdlModelParser";
import "../IoTSolutionItem.scss";

/**
 * Device Model Wizard — Single step:
 * User uploads a DTDL model JSON file, parser extracts capabilities,
 * and user can review/edit the model name and toggle capabilities.
 */
export function DeviceModelStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
}: WizardStepProps) {
  const [fileName, setFileName] = useState<string>(
    wizardContext.modelFileName || ""
  );
  const [modelJson, setModelJson] = useState<unknown>(
    wizardContext.modelJson || null
  );
  const [parseResult, setParseResult] = useState<DtdlParseResult | null>(
    wizardContext.modelParseResult || null
  );
  const [modelName, setModelName] = useState<string>(
    wizardContext.modelName || ""
  );
  const [telemetries, setTelemetries] = useState<DtdlCapability[]>(
    wizardContext.modelTelemetries || []
  );
  const [properties, setProperties] = useState<DtdlCapability[]>(
    wizardContext.modelProperties || []
  );
  const [parseError, setParseError] = useState("");
  const [dragOver, setDragOver] = useState(false);

  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const invalidateGeneratedDataset = useCallback(() => {
    resetStepsFrom(stepIndex + 1);
    updateContext("modeledTablesCreated", false);
    updateContext("propsNormalizedName", "");
    updateContext("propsLkvViewName", "");
    updateContext("modeledDataName", "");
    updateContext("modelCoverageVerified", false);
    updateContext("modelObservedTelemetryFields", []);
    updateContext("modelObservedPropertyFields", []);
  }, [resetStepsFrom, stepIndex, updateContext]);

  const processFile = useCallback(
    (file: File) => {
      setParseError("");
      if (!file.name.endsWith(".json")) {
        setParseError("The selected file is not a valid JSON model file.");
        return;
      }

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const text = e.target?.result as string;
          const parsed = JSON.parse(text);

          // Parse DTDL model
          const result = parseDtdlModel(parsed);

          setFileName(file.name);
          setModelJson(parsed);
          setParseResult(result);
          setModelName(result.displayName);
          setTelemetries(result.telemetries);
          setProperties(result.properties);

          updateContext("modelFileName", file.name);
          updateContext("modelJson", parsed);
          updateContext("modelParseResult", result);
          updateContext("modelName", result.displayName);
          updateContext("modelTelemetries", result.telemetries);
          updateContext("modelProperties", result.properties);
          updateContext("modelUploaded", true);
          invalidateGeneratedDataset();
          invalidateGeneratedDataset();
        } catch (err: unknown) {
          const msg = err instanceof Error ? err.message : "";
          setParseError(
            msg.startsWith("Unsupported model file")
              ? "Unsupported model format. Upload a valid DTDL model or supported device template export."
              : "The model file could not be read. Verify the file format and try again."
          );
          setModelJson(null);
          setParseResult(null);
          updateContext("modelUploaded", false);
        }
      };
      reader.readAsText(file);
    },
    [invalidateGeneratedDataset, updateContext]
  );

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files?.[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => {
    setDragOver(false);
  };

  const toggleTelemetry = (index: number) => {
    const updated = [...telemetries];
    updated[index] = { ...updated[index], included: !updated[index].included };
    setTelemetries(updated);
    updateContext("modelTelemetries", updated);
    invalidateGeneratedDataset();
  };

  const toggleProperty = (index: number) => {
    const updated = [...properties];
    updated[index] = { ...updated[index], included: !updated[index].included };
    setProperties(updated);
    updateContext("modelProperties", updated);
    invalidateGeneratedDataset();
  };

  const setAllTelemetries = (included: boolean) => {
    const updated = telemetries.map((telemetry) => ({ ...telemetry, included }));
    setTelemetries(updated);
    updateContext("modelTelemetries", updated);
    invalidateGeneratedDataset();
  };

  const setAllProperties = (included: boolean) => {
    const updated = properties.map((property) => ({ ...property, included }));
    setProperties(updated);
    updateContext("modelProperties", updated);
    invalidateGeneratedDataset();
  };

  const handleModelNameChange = (value: string) => {
    setModelName(value);
    updateContext("modelName", value);
    invalidateGeneratedDataset();
  };

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Upload Device Model</h2>
      <Text className="iot-solution-step-description" block>
        Upload a DTDL (Digital Twin Definition Language) model to define the
        telemetry and properties that will become typed columns.
      </Text>
      <ul className="iot-solution-outcome-list" style={{ marginTop: "8px" }}>
        <li>
          <strong>Supported model files:</strong> a single DTDL interface, an
          array of related interfaces, or an IoT Central device template export,
          including inherited capabilities and components.
        </li>
        <li>
          <strong>Supported capabilities:</strong> root, inherited, and component
          telemetry and device-reported properties using primitive, enum, and
          complex schemas.
        </li>
        <li>
          <strong>Not included:</strong> commands and cloud-only properties,
          because they are not present in device event data.
        </li>
      </ul>
      <MessageBar intent="info" style={{ marginTop: "12px" }}>
        <MessageBarBody>
          Modeling is optional. Fields that are not included in the model remain
          available in the raw telemetry and property tables for analytics,
          dashboards, and reporting; they will not appear as typed columns in the
          modeled data table.
        </MessageBarBody>
      </MessageBar>

      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".json"
        style={{ display: "none" }}
        onChange={handleFileChange}
      />

      {/* Drop zone */}
      <div
        className={`iot-solution-drop-zone ${dragOver ? "iot-solution-drop-zone--active" : ""} ${modelJson ? "iot-solution-drop-zone--loaded" : ""}`}
        onClick={() => fileInputRef.current?.click()}
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            fileInputRef.current?.click();
          }
        }}
      >
        {modelJson ? (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "8px" }}>
            <DocumentCheckmark24Regular style={{ color: "var(--colorPaletteGreenForeground1)", fontSize: "32px" }} />
            <Text weight="semibold">{fileName}</Text>
            <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
              Click or drop a new file to replace
            </Text>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "8px" }}>
            <ArrowUpload24Regular style={{ fontSize: "32px", color: "var(--colorNeutralForeground3)" }} />
            <Text weight="semibold">Drop DTDL model file here</Text>
            <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
              or click to browse (.json)
            </Text>
          </div>
        )}
      </div>

      {parseError && (
        <MessageBar intent="error" style={{ marginTop: "12px" }}>
          <MessageBarBody>{parseError}</MessageBarBody>
        </MessageBar>
      )}

      {/* Parsed model details */}
      {parseResult && (
        <div style={{ marginTop: "24px" }}>
          {/* Model name input */}
          <div className="iot-solution-field" style={{ maxWidth: "400px", marginBottom: "20px" }}>
            <Label className="iot-solution-field-label" htmlFor="model-name-input">
              Model Name
            </Label>
            <Input
              id="model-name-input"
              value={modelName}
              onChange={(_, data) => handleModelNameChange(data.value)}
            />
          </div>

          <Text className="iot-solution-step-description" block style={{ marginBottom: "12px" }}>
            Select the telemetry and property fields to include in the modeled
            data table. Selected telemetry becomes typed measurement columns, and
            selected properties become current device-state columns added to each
            telemetry event.
          </Text>

          {/* Telemetries */}
          {telemetries.length > 0 && (
            <div className="iot-solution-config-card" style={{ marginBottom: "12px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "12px", marginBottom: "8px" }}>
                <div>
                  <Text weight="semibold" size={400} block>
                    Telemetry Fields
                  </Text>
                  <Text size={200} block style={{ color: "var(--colorNeutralForeground3)" }}>
                    {telemetries.filter((field) => field.included).length} of {telemetries.length} selected
                  </Text>
                </div>
                <Button
                  appearance="subtle"
                  size="small"
                  onClick={() => setAllTelemetries(!telemetries.every((field) => field.included))}
                >
                  {telemetries.every((field) => field.included) ? "Deselect all" : "Select all"}
                </Button>
              </div>
              <div className="iot-solution-field-list">
                {telemetries.map((t, idx) => (
                  <div key={t.name} className="iot-solution-field-list-item">
                    <Checkbox
                      checked={t.included}
                      onChange={() => toggleTelemetry(idx)}
                      label=""
                    />
                    <div style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px" }}>
                      <Text size={300}>{t.displayName || t.name}</Text>
                      <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
                        ({t.name})
                      </Text>
                    </div>
                    <span className="iot-solution-field-badge iot-solution-field-badge--numeric">
                      {t.schema}
                    </span>
                    <Text size={200} style={{ color: "var(--colorNeutralForeground3)", minWidth: "60px", textAlign: "right" }}>
                      → {t.kustoType}
                    </Text>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Properties */}
          {properties.length > 0 && (
            <div className="iot-solution-config-card">
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "12px", marginBottom: "8px" }}>
                <div>
                  <Text weight="semibold" size={400} block>
                    Property Fields
                  </Text>
                  <Text size={200} block style={{ color: "var(--colorNeutralForeground3)" }}>
                    {properties.filter((field) => field.included).length} of {properties.length} selected
                  </Text>
                </div>
                <Button
                  appearance="subtle"
                  size="small"
                  onClick={() => setAllProperties(!properties.every((field) => field.included))}
                >
                  {properties.every((field) => field.included) ? "Deselect all" : "Select all"}
                </Button>
              </div>
              <div className="iot-solution-field-list">
                {properties.map((p, idx) => (
                  <div key={p.name} className="iot-solution-field-list-item">
                    <Checkbox
                      checked={p.included}
                      onChange={() => toggleProperty(idx)}
                      label=""
                    />
                    <div style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px" }}>
                      <Text size={300}>{p.displayName || p.name}</Text>
                      <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
                        ({p.name})
                      </Text>
                    </div>
                    <span className="iot-solution-field-badge iot-solution-field-badge--text">
                      {p.schema}
                    </span>
                    <Text size={200} style={{ color: "var(--colorNeutralForeground3)", minWidth: "60px", textAlign: "right" }}>
                      → {p.kustoType}
                    </Text>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
