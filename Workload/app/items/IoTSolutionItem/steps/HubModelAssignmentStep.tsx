// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import {
  Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle,
  Link, MessageBar, MessageBarBody, Spinner, Text, Tooltip,
} from "@fluentui/react-components";
import { ArrowClockwise20Regular, CheckmarkCircle24Filled, DismissCircle24Filled, Warning24Filled } from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { callNavigationOpenInNewBrowserTab } from "../../../controller/NavigationController";
import { modelValidation, type HubModel } from "../HubModeling.ts";
import type { SourceReadiness } from "../ModelSourceReadiness.ts";
import {
  allAssignmentHubsUnavailable, AssignmentListLoading, assignmentBlocker, assignmentIntentOwner, currentAssignmentHubSummary, currentAssignmentIntent,
  confirmedAssignmentCount, inventoryHasModelAssignments,
  assignmentIssuesForView, assignmentOwner, assignmentReference, ModelAssignmentOperation,
  type AssignmentAttempt, type AssignmentHubSummary, type AssignmentInventory, type AssignmentReceipt, type AssignmentSource, type AssignmentTarget, type AssignmentViewIssue,
} from "../HubAssignment.ts";
import {
  AssignmentApiError, HubAssignmentApi, HubAssignmentReader, assignmentTokenFailure, pollAssignmentReadback,
} from "../HubAssignmentApi.ts";

interface AssignmentProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  getSourceState?: () => SourceReadiness;
  sourceState?: SourceReadiness;
}
type DeviceReadMode = "view" | "refresh" | "more" | "sign-in";
function referenceStatus(target: AssignmentTarget, modelId: string): string {
  return target.reference.kind === "unknown" ? "Assignment not confirmed"
    : target.reference.kind === "unassigned" ? "None"
      : target.reference.ids[0] === modelId ? "Assigned to this model" : "Assigned to another model";
}
function uniqueDevices(targets: AssignmentTarget[]): AssignmentTarget[] {
  const devices = new Map<string, AssignmentTarget>();
  for (const target of targets) if (target.kind === "device" && !devices.has(target.id.toLowerCase())) {
    devices.set(target.id.toLowerCase(), target);
  }
  return [...devices.values()];
}
const order = (target: AssignmentTarget, modelId: string) => target.reference.kind === "unassigned" ? 0
  : target.reference.kind === "unknown" ? 3 : target.reference.ids[0] === modelId ? 1 : 2;

export function AssignmentIssueMessages({ issues, disabled, onRetry }: {
  issues: AssignmentViewIssue[]; disabled: boolean; onRetry: () => void;
}) {
  return <>{issues.map(issue => <MessageBar key={issue.key} intent="warning"><MessageBarBody>
    {issue.message}{issue.count > 1 && <Text size={200}> ({issue.count} checks)</Text>}
    {issue.retryable && <Tooltip content="Refresh devices" relationship="label">
      <Button appearance="subtle" aria-label="Refresh devices" icon={<ArrowClockwise20Regular />}
        disabled={disabled} onClick={onRetry} />
    </Tooltip>}
  </MessageBarBody></MessageBar>)}</>;
}

export function AssignmentReferenceTable({ rows, inventory, modelId, idError, disabled, selectedIds, onToggle, loading = false, refreshing = false }: {
  rows: AssignmentTarget[]; inventory: AssignmentInventory; modelId: string; idError: string; disabled: boolean;
  selectedIds: Set<string>; onToggle: (target: AssignmentTarget, checked: boolean) => void; loading?: boolean; refreshing?: boolean;
}) {
  const devices = refreshing ? [] : rows.filter(target => target.kind === "device");
  return <table className="iot-solution-assignment-table" aria-label="Device model assignments" aria-busy={loading}>
    <thead><tr>
      <th scope="col" id="assignment-device-column">
        <div className="iot-solution-assignment-device-identity">
          <span aria-hidden="true" /><span className="iot-solution-assignment-column-label">Device</span>
        </div>
      </th>
      <th scope="col" id="assignment-model-column"><span className="iot-solution-assignment-column-label">Current model</span></th>
    </tr></thead>
    <tbody>{devices.map(target => {
      const blocker = assignmentBlocker(target) || idError;
      return <tr key={target.id} className="iot-solution-assignment-device">
        <td headers="assignment-device-column"><div className="iot-solution-assignment-device-identity">
          <div className="iot-solution-assignment-device-selection">
          {!blocker && <Checkbox checked={selectedIds.has(target.id.toLowerCase())} disabled={disabled}
            aria-label={`Select ${target.name}`} onChange={(_, data) => onToggle(target, !!data.checked)} />}
          </div><div>
          <Text size={300} block>{target.name}</Text>
          <Text size={200} block className="iot-solution-assignment-device-secondary">
            {target.hubIds.map(id => inventory.hubs.find(hub => hub.id === id)?.name || "Source Hub").join(", ")}
          </Text>
          </div>
        </div></td>
        <td headers="assignment-model-column">
          <Text size={300} block>{referenceStatus(target, modelId)}</Text>
          {target.reference.kind === "assigned" &&
            <Text size={200} block className="iot-solution-assignment-device-secondary">{target.reference.ids[0]}</Text>}
          {blocker && (target.reference.kind !== "assigned" || !!target.writeBlocker) &&
            <Text size={200} block className="iot-solution-assignment-device-secondary">{blocker}</Text>}
        </td>
      </tr>;
    })}</tbody>
  </table>;
}

export function AssignmentHubSummaryBox({ summary, workloadClient, dismissed = false, onDismiss }: {
  workloadClient: WorkloadClientAPI;
  summary: AssignmentHubSummary; dismissed?: boolean; onDismiss?: (event: React.MouseEvent<HTMLButtonElement>) => void;
}) {
  const nonlinked = summary.hubs.filter(hub => hub.state === "unlinked" || hub.state === "pending");
  const unavailable = allAssignmentHubsUnavailable(summary);
  return nonlinked.length ? <>{!dismissed && <MessageBar layout="multiline" bottomReflowSpacer={null} intent="warning"
    className="iot-solution-assignment-hubs iot-solution-assignment-notice" aria-label="Source IoT Hubs" data-summary-kind="warning">
    <MessageBarBody>
      <section className="iot-solution-outcome-section">
        <div className="iot-solution-model-group-header iot-solution-assignment-warning-header">
          <h3 className="iot-solution-outcome-heading">{unavailable ? "Your IoT hubs aren't linked to Azure Device Registry"
            : "Some IoT hubs aren't linked to Azure Device Registry"}</h3>
          {onDismiss && <Button appearance="subtle" size="small" onClick={onDismiss}>Dismiss</Button>}
        </div>
        <Text block>{unavailable ? "Model assignment needs hubs linked to an Azure Device Registry (ADR) namespace."
          : "Devices on these hubs can't be assigned a model until the hubs are linked to an Azure Device Registry (ADR) namespace."}</Text>
        <Text block><Link href="https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal"
          onClick={(event: React.MouseEvent<HTMLAnchorElement>) => {
            event.preventDefault();
            void callNavigationOpenInNewBrowserTab(workloadClient,
              "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal");
          }}
          target="_blank" rel="noopener noreferrer">
          Learn how to link an IoT hub to ADR
        </Link></Text>
        <ul className="iot-solution-outcome-list">{nonlinked.map(hub => <li key={hub.id}>{hub.name}</li>)}</ul>
      </section>
    </MessageBarBody>
  </MessageBar>}
    {unavailable && <Text block data-assignment-unavailable-guidance>
      Select <strong>Continue</strong> to skip this step.
    </Text>}
  </> : null;
}

function AssignmentHubSummaryView(props: AssignmentProps & { owner: string; listLoading: AssignmentListLoading }) {
  const { wizardContext, sourceState, getSourceState, updateContext, workloadClient, owner, listLoading } = props;
  const summary = currentAssignmentHubSummary(wizardContext, sourceState);
  const live = useRef({ wizardContext, sourceState, getSourceState });
  live.current = { wizardContext, sourceState, getSourceState };
  const [attempt, setAttempt] = useState(0);
  const [loading, setLoading] = useState(!summary);
  const [error, setError] = useState("");
  const [needsSignIn, setNeedsSignIn] = useState(false);
  const [dismissedFor, setDismissedFor] = useState("");
  const warningKey = summary ? JSON.stringify([owner, summary.sourceSignature, summary.hubs]) : "";
  useEffect(() => {
    const source: AssignmentSource | undefined = live.current.wizardContext.hubAssignmentSource;
    if (!source || source.key !== (live.current.getSourceState?.() || live.current.sourceState)?.key) {
      listLoading.releaseReservation();
      setLoading(false); setError("The source Hub information is unavailable. Check the source tables, or skip optional assignment.");
      return undefined;
    }
    if (attempt === 0 && currentAssignmentHubSummary(live.current.wizardContext, live.current.sourceState)) return undefined;
    const controller = new AbortController();
    const current = () => {
      const state = live.current.getSourceState?.() || live.current.sourceState;
      return !controller.signal.aborted && assignmentIntentOwner(live.current.wizardContext, state) === owner &&
        currentAssignmentIntent(live.current.wizardContext, state)?.mode === "assign";
    };
    const finish = listLoading.begin();
    setLoading(true); setError(""); setNeedsSignIn(false);
    const api = new HubAssignmentApi({ isCurrent: current, token: async scope => {
      try {
        const result = await workloadClient.auth.acquireFrontendAccessToken({ scopes: [scope] });
        if (typeof result?.token !== "string" || !result.token) throw assignmentTokenFailure(undefined);
        return result.token;
      } catch (failure: unknown) { throw assignmentTokenFailure(failure); }
    } });
    void api.discoverHubs(source, controller.signal).then(value => {
      if (current()) {
        const state = live.current.getSourceState?.() || live.current.sourceState;
        if (value.hubs.some(hub => hub.state === "linked") &&
            !currentAssignmentIntent(live.current.wizardContext, state)?.attempt) listLoading.reserve();
        updateContext("hubAssignmentHubSummary", { owner, value });
      }
    }).catch((failure: unknown) => {
      if (!current()) return;
      setNeedsSignIn(failure instanceof AssignmentApiError && failure.kind === "authentication");
      setError(failure instanceof Error ? failure.message : "The source Hub links could not be checked. Try again.");
    }).finally(() => { finish(); if (current()) setLoading(false); });
    return () => { controller.abort(); finish(); };
  }, [owner, attempt, updateContext, workloadClient, listLoading]);
  const issues = assignmentIssuesForView(summary?.issues || [], "device");
  const retry = () => {
    if (live.current.wizardContext.hubAssignmentBusy ||
        assignmentIntentOwner(live.current.wizardContext, live.current.getSourceState?.() || live.current.sourceState) !== owner) return;
    listLoading.reserve();
    updateContext("hubAssignmentHubSummary", undefined);
    setAttempt(value => value + 1);
  };
  return <>
    {loading && <Spinner size="small" label="Checking source IoT Hub links..." />}
    {summary && <AssignmentHubSummaryBox summary={summary} workloadClient={workloadClient} dismissed={dismissedFor === warningKey} onDismiss={event => {
      const current = live.current;
      const state = current.getSourceState?.() || current.sourceState;
      const currentSummary = currentAssignmentHubSummary(current.wizardContext, state);
      if (assignmentIntentOwner(current.wizardContext, state) !== owner ||
          !currentSummary || JSON.stringify([owner, currentSummary.sourceSignature, currentSummary.hubs]) !== warningKey) return;
      setDismissedFor(warningKey);
      const step = event.currentTarget.closest(".iot-solution-assignment");
      const focus = step?.querySelector<HTMLElement>("#assignment-devices-heading, [aria-label='Device assignment progress'] h3") ||
        step?.closest(".wizard-control")?.querySelector<HTMLElement>('[data-testid="wizard-next-btn"]:not(:disabled)') ||
        step?.querySelector<HTMLElement>(".iot-solution-step-title");
      focus?.focus();
    }} />}
    {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
    {issues.map(issue => <MessageBar key={issue.key} intent="warning"><MessageBarBody>{issue.message}</MessageBarBody></MessageBar>)}
    {(error || issues.length > 0) && <div className="iot-solution-model-actions">
      <Button disabled={loading || !!wizardContext.hubAssignmentBusy} onClick={retry}>Re-check IoT Hubs</Button>
      {(needsSignIn || issues.some(issue => issue.kind === "authentication")) &&
        <Button appearance="primary" disabled={loading || !!wizardContext.hubAssignmentBusy} onClick={retry}>Sign in</Button>}
    </div>}
  </>;
}

export function DpsEnrollmentGuidance({ workloadClient }: { workloadClient: WorkloadClientAPI }) {
  return <section className="iot-solution-outcome-section" aria-label="DPS enrollment guidance" aria-labelledby="dps-enrollment-heading">
      <h3 id="dps-enrollment-heading" className="iot-solution-outcome-heading">Assign the model to new devices automatically</h3>
      <Text block>Update the enrollment groups and individual enrollments in your Device Provisioning Service (DPS) instance. New devices then get this model when they're provisioned, and their data flows into your modeled data.</Text>
      <Text block><Link href="https://aka.ms/iot-hub-agent-ready-data-doc"
        onClick={(event: React.MouseEvent<HTMLAnchorElement>) => {
          event.preventDefault();
          void callNavigationOpenInNewBrowserTab(workloadClient, "https://aka.ms/iot-hub-agent-ready-data-doc");
        }}
        target="_blank" rel="noopener noreferrer">Learn how to update DPS enrollments</Link></Text>
    </section>;
}

function AssignmentDeviceRows({ rows }: { rows: {
  id: string; name: string; state: AssignmentReceipt["state"]; hasError: boolean; message?: string;
}[] }) {
  return <div className="iot-solution-creation-steps">
    {rows.map(row => {
      const done = row.state === "confirmed", warning = row.hasError && row.state === "unknown";
      const labelState = done ? "success" : row.hasError ? warning ? "warning" : "error" : "running";
      return <div key={row.id} className="iot-solution-creation-step" data-assignment-state={row.state}>
        {done ? <CheckmarkCircle24Filled aria-label="Assignment confirmed" primaryFill="var(--colorPaletteGreenForeground1)" />
          : warning ? <Warning24Filled aria-label="Assignment not confirmed" data-assignment-warning primaryFill="var(--colorPaletteYellowForeground1)" />
            : row.hasError ? <DismissCircle24Filled aria-label="Assignment not confirmed" primaryFill="var(--colorPaletteRedForeground1)" />
            : <Spinner size="tiny" aria-label={`Checking assignment for ${row.name}`} />}
        <div className="iot-solution-creation-step-content">
          <Text className={`iot-solution-creation-step-label ${labelState}`}>{row.name}</Text>
          {row.hasError && row.message && <Text size={200}>{row.message}</Text>}
        </div>
      </div>;
    })}
  </div>;
}

function AssignmentProgress({ attempt, writing, history, publish, isCurrent, workloadClient, client, complete, listLoading }: {
  attempt: AssignmentAttempt; writing: boolean; history: AssignmentReceipt[];
  publish: (receipt: AssignmentReceipt) => void; isCurrent: () => boolean;
  workloadClient: WorkloadClientAPI; client?: HubAssignmentApi; complete: boolean;
  listLoading: AssignmentListLoading;
}) {
  const live = useRef({ attempt, history, publish, isCurrent, client });
  live.current = { attempt, history, publish, isCurrent, client };
  const [polling, setPolling] = useState(false);
  const [error, setError] = useState("");
  const [retry, setRetry] = useState(0);
  const heading = useRef<HTMLHeadingElement>(null);
  const receipts = attempt.targetIds.map(id => history.find(receipt => receipt.operationId === attempt.operationId &&
    receipt.targetId.toLowerCase() === id.toLowerCase() && receipt.modelId === attempt.modelId && receipt.kind === "device"));
  const confirmed = receipts.filter(receipt => receipt?.state === "confirmed").length;
  useEffect(() => { heading.current?.focus(); }, [attempt.operationId]);
  useLayoutEffect(() => {
    if (writing) return undefined;
    const controller = new AbortController();
    const operationId = attempt.operationId;
    const current = () => !controller.signal.aborted && live.current.isCurrent() &&
      live.current.attempt.operationId === operationId;
    const snapshot = [...live.current.history];
    for (const id of attempt.targetIds) {
      const index = snapshot.findIndex(receipt => receipt.operationId === operationId && receipt.targetId === id);
      if (index < 0) {
        const receipt: AssignmentReceipt = { operationId, targetId: id, modelId: attempt.modelId, kind: "device",
          name: attempt.targets?.find(target => target.id === id)?.name || id.split("/").at(-1) || id,
          state: "cancelled", message: "No assignment dispatch was recorded for this device. No assignment will be retried." };
        snapshot.push(receipt);
        if (current()) live.current.publish(receipt);
      } else if (snapshot[index].state === "checking" || snapshot[index].state === "assigning") {
        const receipt = { ...snapshot[index], state: snapshot[index].writeDispatched ? "unknown" as const : "cancelled" as const,
          message: snapshot[index].writeDispatched ? "The request may have reached the service. Checking the assignment operation."
            : "Assignment was interrupted before dispatch was confirmed. No assignment will be retried." };
        snapshot[index] = receipt;
        if (current()) live.current.publish(receipt);
      }
    }
    const api = live.current.client || new HubAssignmentApi({ isCurrent: current, token: async scope => {
      try {
        const result = await workloadClient.auth.acquireFrontendAccessToken({ scopes: [scope] });
        if (typeof result?.token !== "string" || !result.token) throw assignmentTokenFailure(undefined);
        return result.token;
      } catch (failure: unknown) { throw assignmentTokenFailure(failure); }
    } });
    const pending = snapshot.some(receipt => receipt.operationId === operationId &&
      ["accepted", "unknown"].includes(receipt.state) && receipt.readbackFailure?.retryable !== false);
    const finish = pending ? listLoading.begin() : (): void => undefined;
    setError("");
    setPolling(pending);
    void pollAssignmentReadback(attempt, snapshot, { current,
      read: (target, modelId, signal, trackingId) => api.readAssignment(target, modelId, signal, trackingId),
      publish: receipt => { if (current()) live.current.publish(receipt); },
    }, controller.signal).catch((failure: unknown) => {
      if (current()) setError(failure instanceof Error ? failure.message : "Assignment verification could not be completed.");
    }).finally(() => {
      finish();
      if (current()) setPolling(false);
    });
    return () => { controller.abort(); finish(); };
  }, [attempt.operationId, writing, retry, workloadClient, listLoading]);
  const retryReadback = () => {
    if (writing || polling || !live.current.isCurrent()) return;
    for (const receipt of receipts) if (receipt?.readbackFailure && receipt.state === "failed") {
      publish({ ...receipt, state: "unknown", readbackFailure: undefined, message: "Checking the same assignment operation again." });
    }
    setRetry(value => value + 1);
  };
  const failed = receipts.some(receipt => receipt && (["failed", "blocked", "cancelled"].includes(receipt.state) || !!receipt.readbackFailure));
  const recoverable = receipts.some(receipt => receipt?.state === "failed" && receipt.readbackFailure);
  const needsSignIn = receipts.some(receipt => receipt?.readbackFailure?.kind === "authentication");
  const active = writing || polling;
  const deviceCount = `${attempt.targetIds.length} ${attempt.targetIds.length === 1 ? "device" : "devices"}`;
  return <section className="iot-solution-model-section" aria-label="Device assignment progress" aria-busy={active}>
    <h3 ref={heading} tabIndex={-1} className="iot-solution-outcome-heading">Assigned devices</h3>
    {active || complete ? <div className="iot-solution-assignment-summary" data-assignment-summary={active ? "active" : "confirmed"}>
      {active ? <Spinner size="tiny" aria-hidden="true" />
        : <CheckmarkCircle24Filled aria-hidden="true" primaryFill="var(--colorPaletteGreenForeground1)" />}
      <Text role="status" className={`iot-solution-assignment-summary-label ${active ? "running" : "success"}`}>
        {active ? `Assigning the model to ${deviceCount}...` : `Model assigned to ${deviceCount}`}
      </Text>
    </div> : <Text role="status" className="iot-solution-model-visually-hidden">{confirmed} of {attempt.targetIds.length} devices confirmed.</Text>}
    {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
    <div className="iot-solution-assignment-device-results">
      <AssignmentDeviceRows rows={attempt.targetIds.map((id, index) => {
        const receipt = receipts[index];
        const stopped = !!receipt && ["failed", "blocked", "cancelled"].includes(receipt.state);
        const hasError = stopped || !!receipt?.readbackFailure;
        const name = receipt?.name || attempt.targets?.find(target => target.id === id)?.name || id.split("/").at(-1) || id;
        return { id, name, state: receipt?.state || "checking", hasError, message: receipt?.message };
        })} />
    </div>
    {failed && <Text block>Not all assignments are confirmed. Review the device results before continuing.</Text>}
    {recoverable && <div className="iot-solution-model-actions">
      {needsSignIn ? <Button disabled={writing || polling} onClick={retryReadback}>Sign in</Button> :
        <Tooltip content="Refresh results" relationship="label"><Button aria-label="Refresh results" icon={<ArrowClockwise20Regular />}
          disabled={writing || polling} onClick={retryReadback} /></Tooltip>}
    </div>}
    {complete && <DpsEnrollmentGuidance workloadClient={workloadClient} />}
  </section>;
}

export function HubModelAssignmentStep(props: AssignmentProps) {
  const { wizardContext, sourceState, getSourceState, updateContext, resetStepsFrom, stepIndex } = props;
  const errors = modelValidation(props.wizardContext.hubModel);
  const model: HubModel | undefined = props.wizardContext.hubModel;
  const heading = useRef<HTMLHeadingElement>(null);
  const pendingSkip = useRef<string>();
  const owner = assignmentIntentOwner(wizardContext, sourceState);
  const intent = currentAssignmentIntent(wizardContext, sourceState);
  const hubs = currentAssignmentHubSummary(wizardContext, sourceState);
  const linked = !!hubs?.hubs.some(hub => hub.state === "linked");
  const live = useRef({ wizardContext, sourceState, getSourceState, resetStepsFrom, stepIndex });
  live.current = { wizardContext, sourceState, getSourceState, resetStepsFrom, stepIndex };
  const pageMounted = useRef(false);
  const listLoading: AssignmentListLoading = useMemo(() => new AssignmentListLoading(owner, () => {
    const current = live.current, state = current.getSourceState?.() || current.sourceState;
    const value = currentAssignmentIntent(current.wizardContext, state);
    if (pageMounted.current && value?.owner === owner && value.listLoading === listLoading) {
      if (listLoading.busy) current.resetStepsFrom(current.stepIndex + 1);
      updateContext("hubAssignmentIntent", { ...value });
    }
  }), [owner, intent?.mode, updateContext]);
  useLayoutEffect(() => {
    pageMounted.current = true;
    listLoading.activate();
    if (intent?.mode === "assign") {
      if (!hubs || linked && !intent.attempt) {
        listLoading.reserve();
        resetStepsFrom(stepIndex + 1);
      }
      if (intent.listLoading !== listLoading) updateContext("hubAssignmentIntent", { ...intent, listLoading });
    }
    return () => { pageMounted.current = false; listLoading.dispose(); };
  }, [owner, listLoading]);
  const busy = !!wizardContext.hubAssignmentBusy;
  useEffect(() => { if (intent?.mode !== "assign" || !linked) heading.current?.focus(); }, [intent?.mode, linked]);
  useLayoutEffect(() => {
    if (wizardContext.hubAssignmentIntent && !intent) {
      resetStepsFrom(stepIndex);
      updateContext("hubAssignmentIntent", undefined);
    }
  }, [wizardContext.hubAssignmentIntent, intent, resetStepsFrom, stepIndex, updateContext]);
  const choose = (mode: "assign" | "skip") => {
    const current = live.current;
    const state = current.getSourceState?.() || current.sourceState;
    if (!owner || assignmentIntentOwner(current.wizardContext, state) !== owner ||
        current.wizardContext.hubAssignmentBusy || mode === "assign" && model?.kind !== "wot") return;
    if (intent?.mode === mode) {
      if (mode === "skip") void props.onNext();
      return;
    }
    pendingSkip.current = mode === "skip" ? owner : undefined;
    if (mode === "assign") {
      const currentHubs = currentAssignmentHubSummary(current.wizardContext, state);
      if (!currentHubs || currentHubs.hubs.some(hub => hub.state === "linked")) listLoading.reserve();
    }
    resetStepsFrom(stepIndex + 1);
    updateContext("hubAssignmentChoices", undefined);
    updateContext("hubAssignmentIntent", { owner, mode, ...(mode === "assign" ? { listLoading } : {}) });
  };
  useEffect(() => {
    const pending = pendingSkip.current;
    if (!pending) return;
    const current = live.current;
    const state = current.getSourceState?.() || current.sourceState;
    if (pending !== owner || assignmentIntentOwner(current.wizardContext, state) !== pending || busy) {
      pendingSkip.current = undefined;
      return;
    }
    if (intent?.mode !== "skip" || intent.owner !== pending) return;
    pendingSkip.current = undefined;
    void props.onNext();
  }, [owner, intent, busy, props.onNext]);
  return <section className="iot-solution-step iot-solution-assignment">
    <h2 ref={heading} tabIndex={-1} className="iot-solution-step-title">Assign the model to devices</h2>
    {errors.length || !model ? <>
      <MessageBar intent="warning"><MessageBarBody>{errors.join(" ")}</MessageBarBody></MessageBar>
      <Button onClick={() => props.onGoToStep("device-model")}>Return to Device Model</Button>
    </> : model.kind === "dtdl" ? <>
      <MessageBar intent="info"><MessageBarBody>
        Model assignment isn't available for DTDL models. To assign a model to devices, go back and upload a WoT Thing Model.
      </MessageBarBody></MessageBar>
      <Text block data-dtdl-assignment-guidance>Select <strong>Continue</strong> to skip this step.</Text>
    </> : <>
      <Text className="iot-solution-step-description" block>
        Optional: assign this model to your devices to get these benefits:
      </Text>
      <ul className="iot-solution-outcome-list">
        <li>Device data is tagged with the model ID automatically</li>
        <li>Data flows into your modeled data without extra setup</li>
        <li>Fewer setup steps as you build your IoT solution</li>
        <li>Easier updates as your devices and data change</li>
      </ul>
      {intent?.mode === "assign" && owner && <>
        <div className="iot-solution-model-task-card" aria-label="Selected device model">
          <dl className="iot-solution-choice-summary">
            <div><dt>Model</dt><dd>{model.title}</dd></div>
            <div><dt>Model ID</dt><dd>{model.id}</dd></div>
          </dl>
        </div>
        <AssignmentHubSummaryView {...props} owner={owner} listLoading={listLoading} />
      </>}
      {intent?.mode !== "assign" && <div className="iot-solution-model-actions">
        <Button appearance="primary" disabled={busy || !owner}
          onClick={() => choose("assign")}>Assign model to devices</Button>
        <Button appearance="secondary" aria-pressed={intent?.mode === "skip"} disabled={busy || !owner}
          onClick={() => choose("skip")}>Skip assignment</Button>
      </div>}
      {!owner && <Text role="status" block>{sourceState?.reason || "Validate the current source tables before choosing an assignment option."}</Text>}
      {intent?.mode === "assign" && linked && <AssignmentContent {...props} model={model} listLoading={listLoading} />}
    </>}
  </section>;
}
function AssignmentContent({ wizardContext, updateContext, workloadClient, getSourceState, sourceState, onGoToStep, model, listLoading }: AssignmentProps & {
  model: HubModel; listLoading: AssignmentListLoading;
}) {
  const source: AssignmentSource | undefined = wizardContext.hubAssignmentSource;
  const live = useRef({ wizardContext, source, model, sourceState, getSourceState });
  live.current = { wizardContext, source, model, sourceState, getSourceState };
  const ownerFor = (state: SourceReadiness | undefined) => {
    const current = live.current;
    if (!state?.ready || !current.source || current.source.key !== state.key || modelValidation(current.model).length ||
        currentAssignmentIntent(current.wizardContext, state)?.mode !== "assign") return "";
    return assignmentOwner(current.source, state.revision, current.model) + `:${current.wizardContext.hubModelRevision || 0}`;
  };
  const ownerNow = () => ownerFor(live.current.getSourceState?.() || live.current.sourceState);
  const owner = ownerFor(sourceState);
  const initial = wizardContext.hubAssignmentInventory;
  const [inventory, setInventory] = useState<AssignmentInventory | undefined>(initial?.owner === owner ? initial.value : undefined);
  const initialChoices = wizardContext.hubAssignmentChoices;
  const [selected, setSelected] = useState<AssignmentTarget[]>(initialChoices?.owner === owner
    ? uniqueDevices(initialChoices.targets) : []);
  const selectedRef = useRef(selected);
  selectedRef.current = selected;
  const [view, setView] = useState("unassigned");
  const [loadingMode, setLoadingMode] = useState<DeviceReadMode>();
  const loading = loadingMode !== undefined;
  const refreshing = loadingMode === "refresh" || loadingMode === "sign-in";
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [needsSignIn, setNeedsSignIn] = useState(false);
  const [confirmation, setConfirmation] = useState(false);
  const [history, setHistory] = useState<AssignmentReceipt[]>(wizardContext.hubAssignmentHistory || []);
  const historyRef = useRef(history);
  const request = useRef<AbortController>();
  const finishListRead = useRef<() => void>();
  const generation = useRef(0);
  const mounted = useRef(false);
  const api = useRef<HubAssignmentApi>();
  const reader = useRef<HubAssignmentReader>();
  const operation = useRef(new ModelAssignmentOperation());
  const targetsHeading = useRef<HTMLHeadingElement>(null);
  const assignButton = useRef<HTMLButtonElement>(null);
  const returnFocus = useRef(false);
  const choose = (targets: AssignmentTarget[]) => {
    const devices = uniqueDevices(targets);
    selectedRef.current = devices;
    setSelected(devices);
    updateContext("hubAssignmentChoices", { owner, targets: devices });
  };
  const publish = (receipt: AssignmentReceipt) => {
    const previous = historyRef.current;
    historyRef.current = [...previous.filter(item => item.operationId !== receipt.operationId || item.targetId !== receipt.targetId), receipt];
    updateContext("hubAssignmentHistory", historyRef.current);
    if (mounted.current) setHistory(historyRef.current);
  };
  const acceptInventory = (value: AssignmentInventory) => {
    const currentInventory = { ...value, targets: uniqueDevices(value.targets) };
    setInventory(currentInventory);
    updateContext("hubAssignmentInventory", { owner, value: currentInventory });
    choose(selectedRef.current.filter(target => {
      const fresh = value.targets.find(item => item.id === target.id);
      return fresh ? !assignmentBlocker(fresh) : value.hasMore;
    }).map(target => value.targets.find(item => item.id === target.id) || target));
  };
  const load = async (mode: DeviceReadMode = "refresh") => {
    const start = ownerNow();
    if (!start || start !== owner || !live.current.source || busy) return;
    request.current?.abort();
    const controller = new AbortController(); request.current = controller;
    const finish = listLoading.begin();
    finishListRead.current?.(); finishListRead.current = finish;
    const ticket = ++generation.current;
    const current = () => mounted.current && !controller.signal.aborted && ticket === generation.current && start === ownerNow();
    setLoadingMode(mode); setError(""); setNeedsSignIn(false);
    try {
      if (!reader.current) {
        const ownerCurrent = () => mounted.current && start === ownerNow();
        const client = new HubAssignmentApi({ isCurrent: ownerCurrent, token: async scope => {
          try {
            const result = await workloadClient.auth.acquireFrontendAccessToken({ scopes: [scope] });
            if (typeof result?.token !== "string" || !result.token) throw assignmentTokenFailure(undefined);
            return result.token;
          } catch (failure: unknown) { throw assignmentTokenFailure(failure); }
        } });
        api.current = client;
        const state = live.current.getSourceState?.() || live.current.sourceState;
        reader.current = new HubAssignmentReader(client, live.current.source, acceptInventory, ownerCurrent, {
          initial: currentAssignmentHubSummary(live.current.wizardContext, state),
          publish: value => {
            if (ownerCurrent()) updateContext("hubAssignmentHubSummary", {
              owner: assignmentIntentOwner(live.current.wizardContext, live.current.getSourceState?.() || live.current.sourceState),
              value,
            });
          },
        });
      }
      await reader.current.load("device", mode === "sign-in" ? "refresh" : mode, controller.signal);
    } catch (failure: unknown) {
      if (!current()) return;
      setNeedsSignIn(failure instanceof AssignmentApiError && failure.kind === "authentication");
      setError(failure instanceof AssignmentApiError && failure.kind === "permission"
        ? "You don't have access to view these references. Ask the service owner for access, then refresh."
        : failure instanceof Error ? failure.message : "Couldn't load devices. Refresh to try again.");
    } finally {
      finish();
      if (finishListRead.current === finish) finishListRead.current = undefined;
      if (current()) setLoadingMode(undefined);
    }
  };
  useEffect(() => { targetsHeading.current?.focus(); }, []);
  useEffect(() => {
    if (!confirmation && returnFocus.current) { returnFocus.current = false; assignButton.current?.focus(); }
  }, [confirmation]);
  useLayoutEffect(() => {
    mounted.current = true;
    setHistory(historyRef.current);
    const choices = initialChoices?.owner === owner
      ? uniqueDevices(initialChoices.targets) : [];
    setInventory(undefined); setSelected(choices);
    selectedRef.current = choices;
    api.current = undefined;
    reader.current = undefined;
    setConfirmation(false); setBusy(false); setError(""); setLoadingMode(undefined); setNeedsSignIn(false);
    if (initialChoices && initialChoices.owner !== owner) updateContext("hubAssignmentChoices", undefined);
    return () => {
      mounted.current = false; generation.current++; request.current?.abort();
      reader.current?.dispose();
      finishListRead.current?.(); finishListRead.current = undefined;
      if (live.current.wizardContext.hubAssignmentBusy === owner) updateContext("hubAssignmentBusy", undefined);
    };
  }, [owner, listLoading]);
  useEffect(() => {
    const state = live.current.getSourceState?.() || live.current.sourceState;
    if (owner && !currentAssignmentIntent(live.current.wizardContext, state)?.attempt) {
      void load("view");
    }
  }, [owner]);
  const modelId = model.id!;
  let idError = "";
  try { assignmentReference(model); } catch (failure: unknown) { idError = failure instanceof Error ? failure.message : "This model reference cannot be assigned."; }
  const devices = uniqueDevices(inventory?.targets || []);
  const rows = devices.filter(target =>
    (view === "all" || target.reference.kind === "unassigned"))
    .sort((left, right) => order(left, modelId) - order(right, modelId) || left.name.localeCompare(right.name));
  const currentIssues = assignmentIssuesForView(inventory?.issues || [], "device");
  const hasMore = inventory?.hasMore;
  const eligible = rows.filter(target => !assignmentBlocker(target) && !idError && !error);
  const selectedIds = new Set(selected.map(target => target.id.toLowerCase()));
  const allShownSelected = eligible.length > 0 && eligible.every(target => selectedIds.has(target.id.toLowerCase()));
  const currentIntent = currentAssignmentIntent(wizardContext, getSourceState?.() || sourceState);
  const assignedCount = confirmedAssignmentCount(currentIntent, history, modelId);
  const settled = !!owner && ownerNow() === owner && !busy && !loading && !error;
  const alreadyAssigned = !currentIntent?.attempt && inventoryHasModelAssignments(inventory);
  const success = settled && (assignedCount > 0 || alreadyAssigned);
  useEffect(() => { if (success && alreadyAssigned) targetsHeading.current?.focus(); }, [success, alreadyAssigned]);
  const closeConfirmation = () => { returnFocus.current = true; setConfirmation(false); };
  const assign = async () => {
    if (busy || loading || !owner || ownerNow() !== owner || !source || !api.current || idError || !selected.length ||
        error || selected.some(target => assignmentBlocker(target))) return;
    const client = api.current, controller = new AbortController();
    request.current?.abort(); request.current = controller;
    const start = owner, targets: AssignmentTarget[] = JSON.parse(JSON.stringify(selected)), operationId = crypto.randomUUID();
    setConfirmation(false); setBusy(true); updateContext("hubAssignmentBusy", start);
    const current = () => mounted.current && ownerNow() === start;
    updateContext("hubAssignmentIntent", { ...currentIntent, attempt: {
      operationId, modelId, targetIds: targets.map(target => target.id), targets,
    } });
    for (const target of targets) publish({ operationId, targetId: target.id, name: target.name, kind: "device", modelId,
      state: "checking", message: "Waiting to assign.", writeDispatched: false });
    try {
      const services = client.assignmentServices(source);
      await operation.current.run(targets, model, true, {
        ...services,
        read: async (target, signal) => {
          const fresh = await services.read(target, signal);
          if (current() && !signal.aborted) reader.current?.updateTarget(fresh);
          return fresh;
        },
      }, current, controller.signal, publish, operationId);
    } catch (failure: unknown) {
      if (current()) setError(failure instanceof Error ? failure.message : "The selected references could not be assigned.");
    } finally {
      if (current()) { setBusy(false); choose([]); updateContext("hubAssignmentBusy", undefined); }
    }
  };
  if (currentIntent?.attempt) return <AssignmentProgress attempt={currentIntent.attempt} writing={busy} history={history}
    listLoading={listLoading}
    complete={success}
    publish={publish} isCurrent={() => ownerNow() === owner &&
      currentAssignmentIntent(live.current.wizardContext, live.current.getSourceState?.() || live.current.sourceState)?.attempt?.operationId === currentIntent.attempt!.operationId}
    workloadClient={workloadClient} client={api.current} />;
  if (success && alreadyAssigned) return <section className="iot-solution-model-section" aria-label="Device model assignments">
    <h3 ref={targetsHeading} tabIndex={-1} className="iot-solution-outcome-heading">Assigned devices</h3>
    <div className="iot-solution-creation-step" role="status" data-all-devices-assigned>
      <CheckmarkCircle24Filled aria-hidden="true" primaryFill="var(--colorPaletteGreenForeground1)" />
      <Text>All devices already have a model assigned</Text>
    </div>
    <DpsEnrollmentGuidance workloadClient={workloadClient} />
  </section>;
  return <section className="iot-solution-model-section" aria-label="Device model assignments">
    <div className="iot-solution-outcome-section">
      <h3 ref={targetsHeading} tabIndex={-1} id="assignment-devices-heading" className="iot-solution-outcome-heading">Select devices</h3>
      <Text block>Choose the devices to assign this model to. Devices that already have a model can't be selected.</Text>
    </div>
    {inventory?.hubs.some(hub => hub.state === "pending" || hub.state === "unverified") && <Text size={200} block>
      Devices from Hubs with pending or unverified ADR links are unavailable. Verified linked Hubs can still have devices.
      You can continue modeling without assignments.
    </Text>}
    {!owner && <MessageBar intent="warning"><MessageBarBody>
      Assignments aren't available for this data source. You can continue modeling without assignments.
      <Button appearance="subtle" onClick={() => onGoToStep("raw-data")}>Review source tables</Button>
    </MessageBarBody></MessageBar>}
    {idError && <MessageBar intent="warning"><MessageBarBody>{idError}</MessageBarBody></MessageBar>}
    <section className="iot-solution-config-card" aria-labelledby="assignment-devices-heading">
    <div className="iot-solution-model-group-header iot-solution-assignment-toolbar" style={{ marginBottom: "var(--spacingVerticalS)" }}>
      <div>
        <Text size={200} block role="status" id="assignment-device-counts" className="iot-solution-assignment-device-secondary">
          {selectedIds.size} selected
        </Text>
      </div>
      <div className="iot-solution-model-actions">
      <Button appearance="subtle" size="small" disabled={!eligible.length || loading || busy}
        title="Select or deselect only eligible unassigned devices in the current view"
        onClick={() => {
          const shownIds = new Set(eligible.map(target => target.id.toLowerCase()));
          choose([...selected.filter(target => !shownIds.has(target.id.toLowerCase())), ...(allShownSelected ? [] : eligible)]);
        }}>{allShownSelected ? "Deselect all" : "Select all"}</Button>
      <Button appearance="subtle" size="small" disabled={loading || busy} aria-describedby="assignment-current-view"
        onClick={() => setView(value => value === "unassigned" ? "all" : "unassigned")}>
        {view === "unassigned" ? "Show all devices" : "Show unassigned devices"}
      </Button>
      <span id="assignment-current-view" role="status" className="iot-solution-model-visually-hidden">
        {view === "unassigned" ? "Current view: unassigned devices." : "Current view: all devices."}
      </span>
      <Tooltip content="Refresh devices" relationship="label">
        <Button aria-label="Refresh devices" icon={<ArrowClockwise20Regular />} disabled={!owner || loading || busy}
          onClick={() => { void load(); }} />
      </Tooltip>
      </div>
    </div>
    {(hasMore || currentIssues.length > 0 || error) && <Text size={200} block className="iot-solution-assignment-device-secondary">
      List incomplete.
    </Text>}
    {loading && <Spinner size="small" label="Loading devices..." />}
    {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
    <AssignmentIssueMessages issues={currentIssues} disabled={loading || busy} onRetry={() => { void load(); }} />
    {(needsSignIn || currentIssues.some(issue => issue.kind === "authentication")) &&
      <Button appearance="primary" disabled={!owner || loading || busy} onClick={() => { void load("sign-in"); }}>Sign in</Button>}
    {inventory && <>
      {!rows.length && !loading && <Text block>{error || currentIssues.length && !inventory.targets.some(target => target.kind === "device")
        ? "Some devices could not be loaded. Refresh to check them."
        : hasMore ? "No matching devices in the loaded results yet. Load more to continue checking."
        : view === "unassigned"
          ? "No confirmed unassigned devices in these results. Use Show all devices to review other assignments."
          : "No devices found for these sources"}</Text>}
      {inventory.targets.some(target => target.kind === "device" && target.reference.kind === "unknown") && <Text block>We couldn't confirm the current assignment. This device can't be changed until it can be checked.</Text>}
      {[...new Set(rows.filter(target => target.reference.kind === "unassigned").map(target => target.writeBlocker).filter(Boolean))]
        .map(blocker => <Text key={blocker} block>{blocker}</Text>)}
      {rows.some(target => target.reference.kind === "unassigned" && target.updatePermission === "unknown" && !target.writeBlocker) &&
        <Text block>Permission to update hasn't been confirmed.</Text>}
      <AssignmentReferenceTable rows={rows} inventory={inventory} modelId={modelId} idError={idError}
        loading={loading} refreshing={refreshing} disabled={busy || loading || !!error} selectedIds={selectedIds} onToggle={(target, checked) => choose(checked
          ? [...selected.filter(item => item.id.toLowerCase() !== target.id.toLowerCase()), target]
          : selected.filter(item => item.id.toLowerCase() !== target.id.toLowerCase()))} />
      {hasMore && <Button disabled={loading || busy} onClick={() => { void load("more"); }}>Load more</Button>}
    </>}
    <div className="iot-solution-model-actions" style={{ marginTop: "var(--spacingVerticalS)" }}>
      <Button ref={assignButton} appearance="primary" disabled={!owner || !selected.length || loading || busy || !!idError ||
        !!error || selected.some(target => !!assignmentBlocker(target))} onClick={() => setConfirmation(true)}>Assign to selected devices</Button>
    </div>
    </section>
    <Dialog open={confirmation} onOpenChange={(_, data) => { if (!data.open && !busy) closeConfirmation(); }}>
      <DialogSurface><DialogBody><DialogTitle>Assign {model.title} to {selected.length} {selected.length === 1 ? "device" : "devices"}?</DialogTitle><DialogContent>
        <Text size={200} block className="iot-solution-assignment-device-secondary">Model ID: {modelId}</Text>
        <ul>{selected.map(target => <li key={target.id}>{target.name}{" \u00b7 "}
          {target.hubIds.map(id => inventory?.hubs.find(hub => hub.id === id)?.name || "Source Hub").join(", ")}
        </li>)}</ul>
      </DialogContent><DialogActions><Button onClick={closeConfirmation}>Cancel</Button>
        <Button appearance="primary" disabled={busy || !owner || ownerNow() !== owner} onClick={() => { void assign(); }}>Assign</Button>
      </DialogActions></DialogBody></DialogSurface>
    </Dialog>
  </section>;
}
