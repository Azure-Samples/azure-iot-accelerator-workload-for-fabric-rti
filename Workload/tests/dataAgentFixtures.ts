// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { fromDtdlModel, fromThingModel, modelingFingerprint, setFieldRole } from "../app/items/IoTSolutionItem/HubModeling.ts";
import { parseDtdlModel } from "../app/items/IoTSolutionItem/DtdlModelParser.ts";
import { resolveThingModel } from "../app/items/IoTSolutionItem/ThingModelParser.ts";
import { buildModeledDataPlan } from "../app/items/IoTSolutionItem/ModeledDataPlan.ts";
import { documentedColumns } from "../app/items/IoTSolutionItem/DataAgentContext.ts";
import { creationRequest } from "./creationMock.ts";
import { metadataDtdlModel, metadataThingModel } from "./fieldSemanticFixtures.ts";

/** Synthetic inspected native docs from the actual model parsers and modeled-data emitter. */
export function dataAgentSourceFixture(kind: "wot" | "dtdl") {
  const document = kind === "wot" ? {
    ...metadataThingModel,
    titles: { "en-US": "WoT fixture device", en: "Fixture device", de: "Testgerät" },
    descriptions: { "en-US": "Synthetic measurements and reported configuration.", de: "Testmessungen." },
  } : metadataDtdlModel.map(model => model["@id"] === "dtmi:fixture:SemanticDevice;2" ? {
    ...model, displayName: { en: "DTDL fixture device", de: "DTDL-Testgerät" },
    description: { en: "Synthetic inherited measurements and reported properties.", de: "Testdaten." },
  } : model);
  const model = kind === "wot" ? fromThingModel(resolveThingModel(document))
    : fromDtdlModel(parseDtdlModel(document), document);
  const request = creationRequest();
  request.names = { ...request.names, modeled: kind === "wot" ? "WoTFixtureModeledData" : "DtdlFixtureModeledData" };
  request.input = { ...request.input, model, fields: model.fields.map(field => field.role ? field
    : setFieldRole(field, field.path[0] === "config" ? "twin" : "telemetry")) };
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
  const output = plan.commands.find(command => command.kind === "table" && command.target === request.names.modeled);
  if (!output || output.kind !== "table") throw new Error("Missing native modeled-table schema in fixture.");
  return {
    database: "Fixture Database", table: request.names.modeled, docstring: plan.documentation.table,
    columns: documentedColumns({ OrderedColumns: output.columns.map(column => ({
      Name: column.name, CslType: column.type, DocString: plan.documentation.columns[column.name],
    })) }),
  };
}
