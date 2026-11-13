import test from "node:test";
import assert from "node:assert/strict";
import { createPlan } from "../src/plan.js";
import { runConfig } from "../src/config.js";
import { ingestionPath, nativeItemUrl } from "../src/paths.js";
import { modelFixture } from "../src/models.js";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { object } from "../src/cloud.js";

const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development", E2E_RETENTION: "retain",
});

test("native item navigation preserves the requested tenant and exact owned workspace/item", () => {
  const plan = createPlan(config);
  const id = "00000000-0000-4000-8000-000000000009";
  for (const kind of ["kustodashboards", "aiskills", "reflexes"] as const) {
    const url = new URL(nativeItemUrl(plan, kind, id));
    assert.equal(url.origin, new URL(config.workspaceUrl).origin);
    assert.equal(url.pathname, `/groups/${config.workspaceId}/${kind}/${id}`);
    assert.equal(url.searchParams.get("ctid"), config.tenantId);
  }
  assert.throws(() => nativeItemUrl(plan, "kustodashboards", "../not-owned"));
});
test("ingestion paths have exact 2/5/3 Hub selections and six distinct raw tables", () => {
  const plan = createPlan(config);
  const paths = ["regular", "combined", "linked"].map(id => ingestionPath(plan, id as "regular" | "combined" | "linked"));
  assert.deepEqual(paths.map(path => path.hubKeys.length), [2, 5, 3]);
  assert.deepEqual(paths[2]?.hubKeys, ["linked1", "linked2", "linked3"]);
  assert.equal(new Set(paths.flatMap(path => [path.telemetryTable, path.propertiesTable])).size, 6);
  assert.equal(new Set(paths.flatMap(path => [path.telemetryStream, path.propertiesStream])).size, 6);
});

test("only nonlinked modeling uses DTDL; both other paths use the example WoT model", async () => {
  for (const path of ["regular", "combined", "linked"] as const) {
    const fixture = modelFixture(path);
    assert.equal(fixture.format, path === "regular" ? "dtdl" : "wot");
    assert.equal(createHash("sha256").update(await readFile(fixture.file)).digest("hex"), fixture.hash);
    assert.ok(fixture.fields.some(field => field.type === "dynamic"));
    assert.ok(fixture.fields.some(field => field.role === "twin"));
  }
});

test("the example DTDL identifiers resolve its bundled component schemas", async () => {
  const fixture = modelFixture("regular");
  let identifiers = 0;
  const bundle: unknown = JSON.parse(await readFile(fixture.file, "utf8"), (key, value: unknown) => {
    if (key === "@id") {
      assert.ok(typeof value === "string");
      assert.match(value, /^dtmi:example:/);
      identifiers++;
    }
    return value;
  });
  assert.ok(Array.isArray(bundle));
  const interfaces = bundle.map((value: unknown) => object(value, "DTDL interface"));
  const interfaceIds = new Set(interfaces.map(model => {
    assert.equal(model["@type"], "Interface");
    assert.ok(typeof model["@id"] === "string");
    return model["@id"];
  }));
  assert.ok(identifiers > interfaces.length);
  const root = interfaces.find(model => model["@id"] === fixture.id);
  assert.ok(root);
  assert.ok(Array.isArray(root.contents));
  const components = root.contents.map((value: unknown) => object(value, "DTDL content"))
    .filter(content => content["@type"] === "Component");
  assert.deepEqual(components.map(component => component.name).sort(), ["powerModule", "thermostat"]);
  for (const component of components) {
    assert.ok(typeof component.schema === "string");
    assert.ok(interfaceIds.has(component.schema));
  }
});
