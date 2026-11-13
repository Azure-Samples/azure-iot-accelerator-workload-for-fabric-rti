// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Restricted, fail-closed replay of emitted dashboard and Data Agent KQL; never a live Kusto client.
export type FixtureRow = Record<string, unknown>;
export const dashboardClock = Date.parse("2026-10-29T12:00:00Z");
/** Offline representation of Kusto's promoted dynamic datetime, retaining all seven fractional digits. */
export class DashboardDateTime {
  readonly iso: string;
  readonly utcKey: string;
  private constructor(iso: string, utcKey: string) { this.iso = iso; this.utcKey = utcKey; }
  static parse(value: string): DashboardDateTime | undefined {
    const match = /^(\d{4}-\d\d-\d\d[T ]\d\d:\d\d:\d\d)(?:\.(\d{1,7}))?(Z|[+-]\d\d:\d\d)?$/.exec(value);
    if (!match) return undefined;
    const [year, month, day, hour, minute, second] = match[1].split(/[-T :]/).map(Number);
    const days = [31, year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (year < 1 || month < 1 || month > 12 || day < 1 || day > days[month - 1] || hour > 23 || minute > 59 || second > 59) return undefined;
    const seconds = Date.parse(match[1] + (match[3] || "Z"));
    if (!Number.isFinite(seconds)) return undefined;
    const utc = new Date(seconds).toISOString().slice(0, 19) + "." + (match[2] || "").padEnd(7, "0") + "Z";
    return new DashboardDateTime(utc, utc);
  }
  toJSON(): string { return this.iso; }
  toString(): string { return this.iso; }
}
const object = (value: unknown): value is FixtureRow => value !== null && typeof value === "object" &&
  !Array.isArray(value) && !(value instanceof DashboardDateTime);
const compareValues = (left: unknown, right: unknown): number => {
  if (left instanceof DashboardDateTime || right instanceof DashboardDateTime) {
    const key = (value: unknown) => value instanceof DashboardDateTime ? value.utcKey
      : typeof value === "number" ? DashboardDateTime.parse(new Date(value).toISOString())?.utcKey : undefined;
    const a = key(left), b = key(right);
    check(a !== undefined && b !== undefined, "Unsupported offline datetime comparison.");
    return a < b ? -1 : a > b ? 1 : 0;
  }
  return typeof left === "number" && typeof right === "number" ? left - right : String(left).localeCompare(String(right));
};
function check(value: unknown, message: string): asserts value { if (!value) throw new Error(message); }
function split(text: string, separator: string): string[] {
  const parts: string[] = [];
  let start = 0, depth = 0, quote = "";
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (quote) { if (char === "\\") index++; else if (char === quote) quote = ""; continue; }
    if (char === '"' || char === "'") quote = char;
    else if ("([{".includes(char)) depth++;
    else if (")]}".includes(char)) depth--;
    else if (!depth && char === separator) { parts.push(text.slice(start, index).trim()); start = index + 1; }
  }
  check(depth === 0 && !quote, "Unbalanced emitted dashboard KQL.");
  parts.push(text.slice(start).trim());
  return parts.filter(Boolean);
}
function literal(token: string): string {
  if (token.startsWith('"')) return JSON.parse(token);
  return token.slice(1, -1).replace(/\\([\s\S])/g, (_match, char: string) => char === "n" ? "\n" : char === "r" ? "\r" : char === "t" ? "\t" : char);
}
function expression(text: string, row: FixtureRow, parameters: FixtureRow, now: number, nativeDynamicDates: boolean): unknown {
  const pattern = /\s+|"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|-?\d+(?:\.\d+)?d?|>=|<=|==|!=|\.\.|![a-zA-Z_]\w*|[a-zA-Z_]\w*|[()[\],.<>+-]/gy;
  const tokens: string[] = [];
  let offset = 0;
  while (offset < text.length) {
    pattern.lastIndex = offset;
    const token = pattern.exec(text);
    check(token && token.index === offset, `Unsupported emitted expression: ${text.slice(offset)}`);
    offset = pattern.lastIndex;
    if (token[0].trim()) tokens.push(token[0]);
  }
  let index = 0;
  const take = (token: string) => check(tokens[index++] === token, `Expected ${token} in ${text}`);
  const type = (value: unknown) => value instanceof DashboardDateTime ? "datetime" : value === null ? "null" : object(value) ? "dictionary" : Array.isArray(value) ? "array"
    : typeof value === "number" ? Number.isInteger(value) ? "long" : "real" : typeof value === "boolean" ? "bool" : typeof value;
  const string = (value: unknown) => value instanceof DashboardDateTime ? value.iso : value === null ? "" : typeof value === "object" ? JSON.stringify(value) : String(value);
  const call = (name: string, args: unknown[]): unknown => {
    switch (name) {
      case "dynamic": return args[0];
      case "tostring": return string(args[0]);
      case "parse_json": {
        if (typeof args[0] !== "string") return args[0];
        try { return JSON.parse(args[0], (_key, value: unknown) => nativeDynamicDates && typeof value === "string" ? DashboardDateTime.parse(value) || value : value); }
        catch { return args[0]; }
      }
      case "todatetime": {
        if (nativeDynamicDates) return args[0] instanceof DashboardDateTime ? args[0] : DashboardDateTime.parse(string(args[0])) || null;
        const date = Date.parse(string(args[0])); return Number.isFinite(date) ? date : null;
      }
      case "todecimal": { const value = Number(args[0]); return args[0] !== null && Number.isFinite(value) ? value : null; }
      case "gettype": return type(args[0]);
      case "isnull": return args[0] === null;
      case "isnotnull": return args[0] !== null;
      case "isempty": return args[0] === null || args[0] === "";
      case "isnotempty": return args[0] !== null && args[0] !== "";
      case "not": return args[0] === null ? null : !args[0];
      case "iff": return args[0] === true ? args[1] : args[2];
      case "case": {
        for (let position = 0; position < args.length - 1; position += 2) if (args[position] === true) return args[position + 1];
        return args[args.length - 1];
      }
      case "coalesce": return args.find(value => value !== null && value !== "") ?? null;
      case "strcat": return args.map(string).join("");
      case "strlen": return string(args[0]).length;
      case "substring": return string(args[0]).substring(Number(args[1]), Number(args[1]) + Number(args[2]));
      case "trim": return string(args[1]).replace(/^\s+|\s+$/g, "");
      case "pack_array": return args;
      case "bag_keys": return object(args[0]) ? Object.keys(args[0]) : null;
      case "bag_has_key": return object(args[0]) && Object.prototype.hasOwnProperty.call(args[0], string(args[1]));
      case "array_index_of": return Array.isArray(args[0]) ? args[0].indexOf(args[1]) : null;
      case "bag_remove_keys": {
        const keys = args[1];
        return object(args[0]) && Array.isArray(keys) ? Object.fromEntries(Object.entries(args[0]).filter(([key]) => !keys.includes(key))) : null;
      }
      case "bag_set_key": return object(args[0]) ? { ...args[0], [string(args[1])]: args[2] } : null;
      case "ingestion_time": return row.ingestedAt ?? now;
      case "ago": return now - Number(args[0]);
      default: throw new Error(`Unsupported dashboard replay function: ${name}`);
    }
  };
  const primary = (): unknown => {
    const token = tokens[index++];
    check(token, `Missing token in ${text}`);
    let value: unknown;
    if (token === "(") { value = or(); take(")"); }
    else if (token === "[") {
      const values: unknown[] = [];
      if (tokens[index] !== "]") do { values.push(or()); } while (tokens[index] === "," && ++index);
      take("]");
      value = values.length === 1 && typeof values[0] === "string" && Object.prototype.hasOwnProperty.call(parameters, values[0])
        ? parameters[values[0]] : values.length === 1 && typeof values[0] === "string" &&
          Object.prototype.hasOwnProperty.call(row, values[0]) ? row[values[0]] : values;
    } else if (token.startsWith('"') || token.startsWith("'")) value = literal(token);
    else if (/^-?\d/.test(token)) value = token.endsWith("d") ? Number(token.slice(0, -1)) * 86400000 : Number(token);
    else if (["true", "false", "null"].includes(token)) value = JSON.parse(token);
    else if (tokens[index] === "(") {
      index++;
      const args: unknown[] = [];
      if (tokens[index] !== ")") do { args.push(or()); } while (tokens[index] === "," && ++index);
      take(")"); value = call(token, args);
    } else {
      check(Object.prototype.hasOwnProperty.call(row, token) || Object.prototype.hasOwnProperty.call(parameters, token),
        `Unknown dashboard replay binding ${token}`);
      value = Object.prototype.hasOwnProperty.call(row, token) ? row[token] : parameters[token];
    }
    while (tokens[index] === "." || tokens[index] === "[") {
      const access = tokens[index++];
      const key = access === "." ? tokens[index++] : or();
      if (access === "[") take("]");
      value = object(value) && typeof key === "string" && Object.prototype.hasOwnProperty.call(value, key) ? value[key] : null;
    }
    return value;
  };
  const sum = (): unknown => {
    let value = primary();
    while (tokens[index] === "+" || tokens[index] === "-") {
      const operator = tokens[index++], right = primary();
      check(typeof value === "number" && typeof right === "number", "Invalid dashboard replay arithmetic.");
      value = operator === "+" ? value + right : value - right;
    }
    return value;
  };
  const comparison = (): unknown => {
    const left = sum(), operator = tokens[index];
    if (operator === "in" || operator === "!in") {
      index++; take("(");
      const values: unknown[] = [];
      do { values.push(or()); } while (tokens[index] === "," && ++index);
      take(")"); return operator === "in" ? values.includes(left) : !values.includes(left);
    }
    if (operator === "between") {
      index++; take("("); const low = primary(); take(".."); const high = primary(); take(")");
      return left !== null && low !== null && high !== null && compareValues(left, low) >= 0 && compareValues(left, high) <= 0;
    }
    if (!["==", "!=", ">=", "<=", ">", "<", "startswith", "!startswith", "endswith_cs"].includes(operator)) return left;
    index++; const right = sum();
    if (operator === "startswith" || operator === "!startswith" || operator === "endswith_cs") {
      const matches = operator === "endswith_cs" ? string(left).endsWith(string(right)) : string(left).startsWith(string(right));
      return operator === "!startswith" ? !matches : matches;
    }
    if (left === null || right === null) return operator === "!=" && left !== right ? true : null;
    if (operator === "==") return left === right;
    if (operator === "!=") return left !== right;
    check(typeof left === "number" && typeof right === "number", `Invalid numeric comparison in ${text}`);
    return operator === ">=" ? left >= right : operator === "<=" ? left <= right : operator === ">" ? left > right : left < right;
  };
  const and = (): unknown => {
    let value = comparison();
    while (tokens[index] === "and") { index++; const right = comparison(); value = value === false || right === false ? false : value === null || right === null ? null : true; }
    return value;
  };
  const or = (): unknown => {
    let value = and();
    while (tokens[index] === "or") { index++; const right = and(); value = value === true || right === true ? true : value === null || right === null ? null : false; }
    return value;
  };
  const value = or();
  check(index === tokens.length, `Unconsumed dashboard KQL: ${tokens.slice(index).join(" ")}`);
  return value;
}

export function replayDashboardQuery(query: string, tables: Record<string, FixtureRow[]>, parameters: FixtureRow = {},
  now = dashboardClock, nativeDynamicDates = false): FixtureRow[] {
  const bindings = new Map<string, FixtureRow[]>();
  const evaluate = (text: string, row: FixtureRow) => expression(text, row, parameters, now, nativeDynamicDates);
  const projectionName = (item: string): string => {
    const alias = /^([A-Za-z_]\w*)\s*=/.exec(item);
    return alias ? alias[1] : item.startsWith("['") && item.endsWith("']") ? literal(item.slice(1, -1)) : item;
  };
  const projected = (text: string, row: FixtureRow): FixtureRow => Object.fromEntries(split(text, ",").map(item => {
    const match = /^([A-Za-z_]\w*)\s*=\s*([\s\S]*)$/.exec(item);
    return [projectionName(item), evaluate(match ? match[2] : item, row)];
  }));
  const pipeline = (text: string): FixtureRow[] => {
    const [source, ...operators] = split(text, "|");
    let rows: FixtureRow[];
    if (source.startsWith("union ")) rows = split(source.slice(6), ",").flatMap(branch => {
      check(branch.startsWith("(") && branch.endsWith(")"), "Expected bounded replay union.");
      return pipeline(branch.slice(1, -1));
    });
    else {
      const table = source.startsWith("['") ? literal(source.slice(1, -1)) : source;
      const value = bindings.get(table) ?? tables[table];
      check(value, `Unknown dashboard replay source: ${table}`);
      rows = value.map(row => ({ ...row }));
    }
    for (const [operatorIndex, operator] of operators.entries()) {
      if (operator.startsWith("extend ")) rows = rows.map(row => ({ ...row, ...projected(operator.slice(7), row) }));
      else if (operator.startsWith("where ")) rows = rows.filter(row => evaluate(operator.slice(6), row) === true);
      else if (operator.startsWith("mv-expand ")) {
        const match = /^mv-expand (\w+) = ([\s\S]*?)(?: to typeof\(string\))?$/.exec(operator);
        check(match, `Unsupported expansion: ${operator}`);
        rows = rows.flatMap(row => {
          const value = evaluate(match[2], row);
          const values = value === null ? [null] : Array.isArray(value) ? value : [value];
          return values.map(item => ({ ...row, [match[1]]: item }));
        });
      } else if (operator.startsWith("project ") || operator.startsWith("distinct ")) {
        const distinct = operator.startsWith("distinct ");
        rows = rows.map(row => projected(operator.slice(distinct ? 9 : 8), row));
        if (distinct) rows = [...new Map(rows.map(row => [JSON.stringify(row), row])).values()];
      } else if (operator.startsWith("summarize ")) {
        const match = /^summarize arg_max\(event_time, (\*|valueType)\) by ([\s\S]*)$/.exec(operator);
        check(match, `Unsupported dashboard aggregation: ${operator}`);
        const groups = new Map<string, FixtureRow>();
        for (const row of rows) {
          const group = projected(match[2], row), key = JSON.stringify(group), previous = groups.get(key);
          if (!previous || compareValues(row.event_time, previous.event_time) > 0) groups.set(key,
            match[1] === "*" ? { ...row, ...group } : { ...group, event_time: row.event_time, valueType: row.valueType });
        }
        rows = [...groups.values()];
      } else if (operator.startsWith("join ")) {
        const match = /^join kind=fullouter \(([\s\S]*)\) on deviceId$/.exec(operator);
        check(match, `Unsupported replay join: ${operator}`);
        const right = pipeline(match[1]);
        const emptyProjection = (stages: string[]): FixtureRow => {
          const projects = stages.filter(stage => stage.startsWith("project "));
          check(projects.length, "Replay join requires explicit projected columns.");
          return Object.fromEntries(split(projects[projects.length - 1].slice(8), ",").map((item): [string, unknown] => {
            const name = projectionName(item), expression = item.slice(item.indexOf("=") + 1).trim();
            // Generated native-string branches use isnotempty; unmatched strings stay empty, not null.
            return [name, name === "deviceId" || stages.some(stage => stage.includes(`isnotempty(${expression})`)) ? "" : null];
          }));
        };
        const emptyLeft = emptyProjection(operators.slice(0, operatorIndex)), emptyRight = emptyProjection(split(match[1], "|"));
        const leftKeys = Object.keys(emptyLeft), rightKeys = Object.keys(emptyRight);
        check(leftKeys.includes("deviceId") && rightKeys.includes("deviceId") &&
          rightKeys.every(key => key === "deviceId" || !leftKeys.includes(key)), "Unsupported replay join column collision.");
        const rightRow = (row?: FixtureRow): FixtureRow => Object.fromEntries(rightKeys.map(key =>
          [key === "deviceId" ? "deviceId1" : key, (row || emptyRight)[key]]));
        const matched = new Set<FixtureRow>(), joined: FixtureRow[] = [];
        for (const left of rows) {
          const matches = right.filter(row => left.deviceId !== null && row.deviceId === left.deviceId);
          if (!matches.length) joined.push({ ...left, ...rightRow() });
          for (const row of matches) { matched.add(row); joined.push({ ...left, ...rightRow(row) }); }
        }
        for (const row of right) if (!matched.has(row)) joined.push({ ...emptyLeft, ...rightRow(row) });
        rows = joined;
      } else if (operator.startsWith("order by ")) {
        const match = /^order by (\w+) (asc|desc)$/.exec(operator);
        check(match, `Unsupported ordering: ${operator}`);
        rows.sort((a, b) => {
          const left = a[match[1]], right = b[match[1]];
          const order = compareValues(left, right);
          return match[2] === "asc" ? order : -order;
        });
      } else if (/^(limit|take) \d+$/.test(operator)) rows = rows.slice(0, Number(operator.split(" ")[1]));
      else throw new Error(`Unsupported emitted dashboard operator: ${operator}`);
    }
    return rows;
  };
  const statements = split(query, ";");
  for (const statement of statements.slice(0, -1)) {
    const match = /^let (\w+) = ([\s\S]*)$/.exec(statement);
    check(match, `Unsupported dashboard declaration: ${statement}`);
    const grouped = /^\(([\s\S]*)\)$/.exec(match[2]);
    bindings.set(match[1], pipeline(grouped ? grouped[1] : match[2]));
  }
  return pipeline(statements[statements.length - 1]);
}

export function dashboardRow(value: unknown, options: {
  do?: boolean; stream?: "telemetry" | "twin"; component?: string; module?: string; profile?: "native" | "cloud" | "empty-native";
  time?: string; device?: string;
} = {}): FixtureRow {
  const { component = "", module = "", stream = "telemetry", profile = "cloud", time = "2026-10-29T11:00:00Z", device = "device-1" } = options;
  const identity = module ? `${device}/${module}` : device;
  const data = options.do === false ? value : {
    schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
    entity_id: "not-a-device-id", event_kind: stream === "telemetry" ? "measurement" : "lifecycle", event_time: time,
    payload: { value },
  };
  return { data: JSON.stringify(data),
    headers: profile === "native" ? { IoTConnectionDeviceId: device, IoTConnectionModuleId: module, IoTSubject: component, IoTEnqueueTime: "0001-01-01T00:00:00Z" }
      : profile === "empty-native" ? { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" } : {},
    user_headers: profile === "native" ? {} : { cloudEvents_subject: identity, cloudEvents_deviceid: device,
      cloudEvents_moduleid: module, cloudEvents_iothubdtsubject: component, cloudEvents_time: time },
    ingestedAt: dashboardClock - 60000,
  };
}
