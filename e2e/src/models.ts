import { join } from "node:path";
import { root } from "./config.js";
import type { PathId } from "./paths.js";

export interface SelectedField {
  id: string;
  column: string;
  role: "telemetry" | "twin";
  type: "real" | "string" | "dynamic";
  component?: string;
  property: string;
}

export function modelFixture(path: PathId) {
  const dtdl = path === "regular";
  const fields: SelectedField[] = dtdl ? [
    { id: "/root/Telemetry/temperature", column: "temperature", role: "telemetry", type: "real", property: "temperature" },
    { id: "/root/Telemetry/gpsLocation", column: "gpsLocation", role: "telemetry", type: "dynamic", property: "gpsLocation" },
    { id: "/components/thermostat/Telemetry/targetTemperature", column: "thermostat_targetTemperature", role: "telemetry", type: "real", component: "thermostat", property: "targetTemperature" },
    { id: "/components/powerModule/Telemetry/powerConsumption", column: "powerModule_powerConsumption", role: "telemetry", type: "real", component: "powerModule", property: "powerConsumption" },
    { id: "/root/Property/firmwareVersion", column: "firmwareVersion", role: "twin", type: "string", property: "firmwareVersion" },
    { id: "/root/Property/maintenanceSchedule", column: "maintenanceSchedule", role: "twin", type: "dynamic", property: "maintenanceSchedule" },
    { id: "/components/thermostat/Property/setpoint", column: "thermostat_setpoint", role: "twin", type: "real", component: "thermostat", property: "setpoint" },
    { id: "/components/powerModule/Property/batteryHealth", column: "powerModule_batteryHealth", role: "twin", type: "string", component: "powerModule", property: "batteryHealth" },
  ] : [
    { id: "/properties/temperature", column: "temperature", role: "telemetry", type: "real", property: "temperature" },
    { id: "/properties/gpsLocation", column: "gpsLocation", role: "telemetry", type: "dynamic", property: "gpsLocation" },
    { id: "/properties/firmwareVersion", column: "firmwareVersion", role: "twin", type: "string", property: "firmwareVersion" },
    { id: "/properties/tags", column: "tags", role: "twin", type: "dynamic", property: "tags" },
    { id: "/properties/maintenanceSchedule", column: "maintenanceSchedule", role: "twin", type: "dynamic", property: "maintenanceSchedule" },
  ];
  return {
    format: dtdl ? "dtdl" : "wot",
    title: "ThermoPro X200",
    id: dtdl ? "dtmi:example:ThermoPro_X200;1" : "urn:example:ThermoProX200",
    file: join(root, "fixtures", dtdl ? "ThermoProX200-dtdl.json" : "ThermoProX200-wot-tm.json"),
    hash: dtdl ? "9b1d59c20f62401fc547eb3b292a28734021fe112733d50197b22524d922bd1b" : "169b0a51ea472bb3ad60256f587df0349bdb16bbbb55d52885064516c935a676",
    fields,
  };
}

export type ModelFixture = ReturnType<typeof modelFixture>;
