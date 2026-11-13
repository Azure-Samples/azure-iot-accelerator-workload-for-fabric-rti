// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { canonicalJson, type SampleStream } from "./ModelAuthoring.ts";
import { HUB_METADATA } from "./HubTransport.ts";
import {
  COVERAGE_SAMPLING_CONTRACT, coverageReady, effectiveScope, inspectFieldEvidence, kqlString, modelingFingerprint, nativeBindingIssues, processingBindingSamples, RAW_SCOPE_CONTRACT, verifiedEmptyCoverage,
  type CoverageSnapshot, type ModelField, type ModelingInput, type ProcessingScope, type inspectStreamCoverage,
} from "./HubModeling.ts";

export function coverageConditionExample(fields: ModelField[], stream: SampleStream, multiline = false): string {
  const field = fields.find(item => item.included && item.role === stream);
  const eh = HUB_METADATA.eventhub, ce = HUB_METADATA.cloudEvents;
  const andJoin = multiline ? "\n\tand " : " and ";
  const orJoin = multiline ? "\n\tor " : " or ";
  if (!field) return [`isnotempty(tostring(headers[${kqlString(eh.device)}]))`,
    `isnotempty(tostring(user_headers[${kqlString(stream === "telemetry" ? ce.subject : ce.device)}]))`].join(orJoin);
  const access = (path: string[]) => ["payload", "value", ...(stream === "twin" ? ["properties", "reported"] : []), ...path]
    .reduce((expression, key) => /^[A-Za-z_]\w*$/.test(key) ? `${expression}.${key}` : `${expression}[${kqlString(key)}]`, "data");
  const value = access(stream === "twin" && field.component ? [field.component, ...field.path] : field.path);
  const condition = `isnotnull(${value})`;
  if (stream === "telemetry") {
    const components = [`tostring(headers[${kqlString(eh.component)}])`, `tostring(user_headers[${kqlString(ce.component)}])`];
    const componentName = field.component;
    return componentName
      ? `(${components.map(component => `${component} == ${kqlString(componentName)}`).join(orJoin)})${andJoin}${condition}`
      : `${components.map(component => `isempty(${component})`).join(andJoin)}${andJoin}${condition}`;
  }
  if (field.component) return `tostring(${access([field.component, "__t"])}) == "c"${andJoin}${condition}`;
  return fields.some(item => item.included && item.component)
    ? `tostring(${access([field.path[0], "__t"])}) != "c"${andJoin}${condition}` : condition;
}

export interface StreamTicket { base: string; revision: number; stream: SampleStream; generation: number; predicate: string }
type StreamResult = Awaited<ReturnType<typeof inspectStreamCoverage>>;
interface StreamCheck {
  generation: number; phase: "unchecked" | "checking" | "checked" | "error"; edited?: boolean;
  proof?: { ticket: StreamTicket; result: StreamResult }; error?: string;
}
export interface StreamChecks {
  base: string; revision: number; predicates: { telemetry: string; twin: string }; telemetry: StreamCheck; twin: StreamCheck;
}
function baseKey(input: ModelingInput, sourceRevision: number): string {
  return canonicalJson({ predicateContract: RAW_SCOPE_CONTRACT, timeColumn: "event_time", coverageSampling: COVERAGE_SAMPLING_CONTRACT, source: input.source, sourceRevision, model: input.model, fields: input.fields, limit: input.limit,
    mode: input.scope.mode, additional: input.scope.additional, operator: input.scope.operator, reference: input.scope.reference });
}
export function currentStreamChecks(input: ModelingInput, sourceRevision: number, previous?: StreamChecks): StreamChecks {
  const base = baseKey(input, sourceRevision);
  const predicates = { telemetry: input.scope.telemetry, twin: input.scope.twin };
  if (previous?.base === base) {
    if (previous.predicates.telemetry === predicates.telemetry && previous.predicates.twin === predicates.twin) return previous;
    return { ...previous, predicates, ...Object.fromEntries((["telemetry", "twin"] as const)
      .filter(stream => previous.predicates[stream] !== predicates[stream])
      .map(stream => [stream, { generation: previous[stream].generation + 1, phase: "unchecked", edited: true }])) };
  }
  return { base, revision: (previous?.revision || 0) + 1, predicates,
    telemetry: { generation: (previous?.telemetry.generation || 0) + 1, phase: "unchecked" },
    twin: { generation: (previous?.twin.generation || 0) + 1, phase: "unchecked" } };
}
export function editStreamChecks(input: ModelingInput, scope: ProcessingScope, sourceRevision: number, previous?: StreamChecks): StreamChecks {
  return currentStreamChecks({ ...input, scope }, sourceRevision, previous);
}
export function beginStreamCheck(input: ModelingInput, sourceRevision: number, stream: SampleStream, previous?: StreamChecks) {
  const next = currentStreamChecks(input, sourceRevision, previous);
  const generation = next[stream].generation + 1;
  const ticket: StreamTicket = { base: next.base, revision: next.revision, stream, generation, predicate: effectiveScope(input.scope, stream) };
  return { checks: { ...next, [stream]: { generation, phase: "checking" as const } }, ticket };
}
function accepts(input: ModelingInput, checks: StreamChecks, ticket: StreamTicket): boolean {
  return checks.base === ticket.base && checks.revision === ticket.revision &&
    checks[ticket.stream].generation === ticket.generation && effectiveScope(input.scope, ticket.stream) === ticket.predicate;
}
export function currentStreamResult(input: ModelingInput, sourceRevision: number, checks: StreamChecks | undefined,
  stream: SampleStream): StreamResult | undefined {
  if (input.scope.mode === "unselected" || input.scope.additional === "unselected" || !checks || checks.base !== baseKey(input, sourceRevision)) return undefined;
  const check = checks[stream], proof = check.proof;
  return check.phase === "checked" && proof?.ticket.stream === stream && accepts(input, checks, proof.ticket)
    ? proof.result : undefined;
}
export function finishStreamCheck(input: ModelingInput, checks: StreamChecks, ticket: StreamTicket,
  outcome: { result: StreamResult } | { error: string }): StreamChecks {
  if (checks[ticket.stream].phase !== "checking" || !accepts(input, checks, ticket)) return checks;
  return { ...checks, [ticket.stream]: "result" in outcome
    ? { generation: ticket.generation, phase: "checked", proof: { ticket, result: outcome.result } }
    : { generation: ticket.generation, phase: "error", error: outcome.error } };
}
export function cancelStreamCheck(checks: StreamChecks, ticket: StreamTicket): StreamChecks {
  const current = checks[ticket.stream];
  if (current.phase !== "checking" || checks.base !== ticket.base || checks.revision !== ticket.revision ||
      current.generation !== ticket.generation) return checks;
  return { ...checks, [ticket.stream]: { generation: current.generation + 1, phase: "unchecked", edited: true } };
}
export function assembleStreamCoverage(input: ModelingInput, sourceRevision: number, checks?: StreamChecks): CoverageSnapshot | undefined {
  if (input.scope.mode === "unselected" || input.scope.additional === "unselected" || !checks || checks.base !== baseKey(input, sourceRevision)) return undefined;
  const telemetry = checks.telemetry.proof, twin = checks.twin.proof;
  if (!telemetry || !twin || telemetry.ticket.stream !== "telemetry" || twin.ticket.stream !== "twin" ||
      checks.telemetry.phase !== "checked" || checks.twin.phase !== "checked" ||
      !accepts(input, checks, telemetry.ticket) || !accepts(input, checks, twin.ticket)) return undefined;
  return { fingerprint: modelingFingerprint(input), kind: "processing",
    telemetry: telemetry.result.sample, twin: twin.result.sample,
    sourceProbe: { telemetry: telemetry.result.probe, twin: twin.result.probe },
    fields: inspectFieldEvidence(input.fields, telemetry.result.sample, twin.result.sample),
    nativeIssues: nativeBindingIssues(telemetry.result.sample, twin.result.sample),
    checkedAt: [telemetry.result.checkedAt, twin.result.checkedAt].sort().at(-1)! };
}
export function streamCoverageReady(input: ModelingInput, sourceRevision: number, checks: StreamChecks | undefined, waiting: boolean): boolean {
  const snapshot = assembleStreamCoverage(input, sourceRevision, checks);
  if (!snapshot || !coverageReady(input, snapshot, waiting)) return false;
  const bindings = processingBindingSamples(snapshot);
  return nativeBindingIssues(bindings.telemetry, bindings.twin).length === 0;
}
export function modelCoverageReady(input: ModelingInput, sourceRevision: number, checks: StreamChecks | undefined): boolean {
  const snapshot = assembleStreamCoverage(input, sourceRevision, checks);
  return streamCoverageReady(input, sourceRevision, checks, verifiedEmptyCoverage(input, snapshot));
}
