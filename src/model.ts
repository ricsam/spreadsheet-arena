import { FormulaEngine, type WorkbookData } from "@ricsam/formula-engine";
import { spreadsheetToFormulaEngineDetailed } from "@ricsam/excel-to-formula-engine";

export type Pane = {
  id: string;
  workbookName: string;
  sheetName: string;
  cell?: string;
  zoom?: number;
};
export type Layout =
  | { kind: "pane"; pane: Pane }
  | {
      kind: "split";
      id: string;
      direction: "horizontal" | "vertical";
      ratio: number;
      first: Layout;
      second: Layout;
    };
export interface WorkspaceSnapshot {
  format: "spreadsheet-arena";
  version: 1;
  engine: string;
  layout: Layout | null;
  activePaneId: string | null;
  view?: { sidebar: boolean; showFormulas: boolean; maximized: string | null };
}

export type WorkspaceView = Pick<
  WorkspaceSnapshot,
  "layout" | "activePaneId" | "view"
>;
export type LoadedWorkspace = {
  snapshot: WorkspaceSnapshot;
  engine: FormulaEngine;
};

export const MAX_IMPORT_FILE_BYTES = 25 * 1024 * 1024;
// JSON backups are uncompressed and substantially larger than their source XLSX.
export const MAX_WORKSPACE_FILE_BYTES = 256 * 1024 * 1024;
export const MAX_IMPORT_FILES = 20;

export function uniqueWorkbookName(
  engine: FormulaEngine,
  desired: string,
): string {
  const base = desired.trim() || "Untitled workbook";
  const names = new Set(
    [...engine.getWorkbooks().keys()].map((n) => n.toLowerCase()),
  );
  let name = base;
  for (let i = 2; names.has(name.toLowerCase()); i++) name = `${base} (${i})`;
  return name;
}
export function createEmptyWorkbook(
  engine: FormulaEngine,
  desired: string,
): string {
  const workbookName = uniqueWorkbookName(engine, desired);
  engine.addWorkbook({ workbookName, data: { sheets: [{ name: "Sheet1" }] } });
  return workbookName;
}
export function createDemoEngine(): FormulaEngine {
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: "Assumptions",
    data: {
      sheets: [
        {
          name: "Inputs",
          sheetMetadata: {
            arenaColumnWidths: { A: 185, B: 100 },
            arenaRowHeights: { "1": 38, "8": 64 },
          },
          content: {
            A1: "LAUNCH INPUTS",
            A3: "Input",
            B3: "Value",
            A4: "Cost per guest",
            B4: 45,
            A5: "Venue hire",
            B5: 2400,
            A6: "Contingency",
            B6: 0.1,
            A8: "Edit these inputs to update the launch budget.",
          },
        },
      ],
    },
  });
  engine.addWorkbook({
    workbookName: "Launch budget",
    data: {
      sheets: [
        {
          name: "Overview",
          sheetMetadata: {
            arenaColumnWidths: { A: 240, B: 80, C: 100, D: 115 },
            arenaRowHeights: { "1": 38, "2": 48, "12": 34 },
          },
          content: {
            A1: "A SMALL LAUNCH.",
            A2: "A live budget connected to Assumptions → Inputs",
            A4: "Line item",
            B4: "Quantity",
            C4: "Unit cost",
            D4: "Budget",
            A5: "Guests & catering",
            B5: 80,
            C5: "=[Assumptions]Inputs!B4",
            D5: "=B5*C5",
            A6: "Venue",
            B6: 1,
            C6: "=[Assumptions]Inputs!B5",
            D6: "=B6*C6",
            A7: "Creative & signage",
            B7: 1,
            C7: 850,
            D7: "=B7*C7",
            A9: "Subtotal",
            D9: "=SUM(D5:D7)",
            A10: "Contingency",
            D10: "=D9*[Assumptions]Inputs!B6",
            A12: "TOTAL LAUNCH BUDGET",
            D12: "=SUM(D9:D10)",
          },
        },
      ],
    },
  });
  for (const [workbookName, sheetName, headerRow] of [
    ["Launch budget", "Overview", 3],
    ["Assumptions", "Inputs", 2],
  ] as const) {
    engine.addCellStyle({
      areas: [
        {
          workbookName,
          sheetName,
          range: {
            start: { row: 0, col: 0 },
            end: {
              row: { type: "number", value: 0 },
              col: { type: "number", value: 3 },
            },
          },
        },
      ],
      style: {
        backgroundColor: "#172f2b",
        color: "#ffffff",
        bold: true,
        fontSize: 14,
        wrapText: true,
      },
    });
    engine.addCellStyle({
      areas: [
        {
          workbookName,
          sheetName,
          range: {
            start: { row: 1, col: 0 },
            end: {
              row: { type: "number", value: 7 },
              col: { type: "number", value: 0 },
            },
          },
        },
      ],
      style: { wrapText: true },
    });
    engine.addCellStyle({
      areas: [
        {
          workbookName,
          sheetName,
          range: {
            start: { row: headerRow, col: 0 },
            end: {
              row: { type: "number", value: headerRow },
              col: { type: "number", value: 3 },
            },
          },
        },
      ],
      style: { backgroundColor: "#e2eee7", bold: true, color: "#244b3e" },
    });
  }
  engine.clearUndoRedoHistory();
  return engine;
}

export function decodeEngine(serialized: string): FormulaEngine {
  try {
    if (typeof serialized !== "string")
      throw new Error("Expected serialized JSON text");
    const engine = new FormulaEngine();
    engine.resetToSerializedEngine(serialized);
    for (const [name, workbook] of engine.getWorkbooks()) {
      if (
        !name ||
        workbook.name !== name ||
        !(workbook.sheets instanceof Map) ||
        !workbook.sheets.size
      )
        throw new Error("Invalid workbook or missing sheets");
      for (const [sheetName, sheet] of workbook.sheets) {
        if (
          !sheetName ||
          sheet.name !== sheetName ||
          !(sheet.content instanceof Map)
        )
          throw new Error("Invalid sheet");
      }
    }
    return engine;
  } catch (error) {
    throw new Error(`Cannot read spreadsheet engine: ${message(error)}`);
  }
}
const message = (error: unknown) =>
  error instanceof Error ? error.message : String(error);

// Native snapshots use tagged Maps. Keep the engine's own representation rather
// than reconstructing WorkbookData (which would drop style/metadata/table state).
type Native = ReturnType<typeof JSON.parse>;
const native = (engine: FormulaEngine): Native =>
  JSON.parse(engine.serializeEngine({ includeEvaluationState: false }));
function clearDerived(snapshot: Native) {
  snapshot.managers.dependency = { nodes: [], spilledValues: [] };
  snapshot.managers.cache = { evaluationOrders: [], sccs: [] };
}
export function exportWorkbook(engine: FormulaEngine, name: string): string {
  if (!engine.hasWorkbook(name)) throw new Error(`Workbook not found: ${name}`);
  const snapshot = native(engine),
    m = snapshot.managers;
  for (const map of [
    m.workbook,
    m.table,
    m.namedExpression.sheetExpressions,
    m.namedExpression.workbookExpressions,
  ]) {
    map.entries = map.entries.filter(
      ([key]: [string, unknown]) => key === name,
    );
  }
  for (const key of ["conditionalStyles", "cellStyles", "cellDataTypes"]) {
    m.style[key] = m.style[key]
      .map((item: Native) => ({
        ...item,
        areas: item.areas.filter((a: Native) => a.workbookName === name),
      }))
      .filter((item: Native) => item.areas.length);
  }
  m.rangeMetadata = m.rangeMetadata
    .map((item: Native) => ({
      ...item,
      areas: item.areas.filter((a: Native) => a.workbookName === name),
    }))
    .filter((item: Native) => item.areas.length);
  // Include only global definitions reachable from this workbook, not unrelated
  // workspace-level constants. Walk named definitions transitively as well.
  const neededGlobals = new Set<string>();
  const visitedNames = new Set<string>();
  const state = engine.getState().namedExpressions;
  function collect(formula: string, sheetName: string) {
    for (const reference of engine.analyzeFormula({
      formula,
      origin: { workbookName: name, sheetName, rowIndex: 0, colIndex: 0 },
    }).references) {
      if (reference.resolution.status !== "resolved") continue;
      for (const target of reference.resolution.targets) {
        if (target.type !== "named-expression") continue;
        const scope = target.scope;
        const key = JSON.stringify([scope, target.name]);
        if (visitedNames.has(key)) continue;
        visitedNames.add(key);
        const definitions =
          scope.type === "global"
            ? state.globalExpressions
            : scope.type === "workbook"
              ? state.workbookExpressions.get(scope.workbookName)
              : state.sheetExpressions
                  .get(scope.workbookName)
                  ?.get(scope.sheetName);
        const match = [...(definitions?.entries() ?? [])].find(
          ([n]) => n.toLowerCase() === target.name.toLowerCase(),
        );
        if (!match) continue;
        if (scope.type === "global") neededGlobals.add(match[0]);
        collect(
          match[1].expression,
          scope.type === "sheet" ? scope.sheetName : sheetName,
        );
      }
    }
  }
  for (const [sheetName, sheet] of engine.getWorkbooks().get(name)!.sheets) {
    for (const content of sheet.content.values())
      if (typeof content === "string" && content.startsWith("="))
        collect(content, sheetName);
    for (const expression of state.sheetExpressions
      .get(name)
      ?.get(sheetName)
      ?.values() ?? [])
      collect(expression.expression, sheetName);
  }
  const firstSheet = engine
    .getWorkbooks()
    .get(name)!
    .sheets.keys()
    .next().value!;
  for (const expression of state.workbookExpressions.get(name)?.values() ?? [])
    collect(expression.expression, firstSheet);
  m.namedExpression.globalExpressions.entries =
    m.namedExpression.globalExpressions.entries.filter(
      ([key]: [string, unknown]) => neededGlobals.has(key),
    );
  // Tracked references are scoped by address; never export other books' ranges.
  m.reference.entries = m.reference.entries.filter(
    ([, ref]: [string, Native]) => ref.address.workbookName === name,
  );
  clearDerived(snapshot);
  return decodeEngine(JSON.stringify(snapshot)).serializeEngine({
    includeEvaluationState: false,
  });
}
function mergeEngines(
  existing: FormulaEngine,
  incoming: FormulaEngine,
  warnings: string[],
): string {
  const a = native(existing),
    b = native(incoming),
    x = a.managers,
    y = b.managers;
  for (const key of ["workbook", "table"])
    x[key].entries.push(...y[key].entries);
  // Reference handles are consumer-owned runtime IDs. Retain existing handles;
  // allocate new IDs when importing a second copy of the same native snapshot.
  const referenceIds = new Set(
    x.reference.entries.map(([id]: [string, unknown]) => id),
  );
  for (const [oldId, reference] of y.reference.entries) {
    const id = referenceIds.has(oldId) ? crypto.randomUUID() : oldId;
    referenceIds.add(id);
    x.reference.entries.push([id, { ...reference, id }]);
  }
  const metadataIds = new Set(x.rangeMetadata.map((item: Native) => item.id));
  for (const item of y.rangeMetadata) {
    if (metadataIds.has(item.id)) item.id = crypto.randomUUID();
    metadataIds.add(item.id);
  }
  for (const key of ["sheetExpressions", "workbookExpressions"])
    x.namedExpression[key].entries.push(...y.namedExpression[key].entries);
  // Global names belong to the workspace: preserve existing definitions on conflict.
  const globals = new Set(
    x.namedExpression.globalExpressions.entries.map(([k]: [string, unknown]) =>
      k.toLowerCase(),
    ),
  );
  for (const [key, expression] of y.namedExpression.globalExpressions.entries) {
    if (globals.has(key.toLowerCase()))
      warnings.push(
        `Global named expression “${key}” already exists; kept the existing definition. Imported formulas using it may calculate differently.`,
      );
    else {
      x.namedExpression.globalExpressions.entries.push([key, expression]);
      globals.add(key.toLowerCase());
    }
  }
  for (const key of ["conditionalStyles", "cellStyles", "cellDataTypes"])
    x.style[key].push(...y.style[key]);
  x.rangeMetadata.push(...y.rangeMetadata);
  clearDerived(a);
  return decodeEngine(JSON.stringify(a)).serializeEngine();
}
export async function importFiles(
  engine: FormulaEngine,
  files: File[],
): Promise<{ names: string[]; warnings: string[]; errors: string[] }> {
  const result = {
    names: [] as string[],
    warnings: [] as string[],
    errors: [] as string[],
  };
  if (files.length > MAX_IMPORT_FILES)
    return {
      ...result,
      errors: [
        `Import at most ${MAX_IMPORT_FILES} files at a time; no files were imported.`,
      ],
    };
  for (const file of files) {
    try {
      if (file.size > MAX_IMPORT_FILE_BYTES)
        throw new Error("File exceeds the 25 MB import limit");
      let incoming: FormulaEngine;
      const warnings: string[] = [];
      if (/\.json$/i.test(file.name)) {
        const raw: unknown = JSON.parse(await file.text());
        if (isRecord(raw) && raw.format === "spreadsheet-arena")
          incoming = decodeEngine(validateSnapshot(raw).engine);
        else if (isRecord(raw) && Array.isArray(raw.sheets)) {
          if (
            !raw.sheets.length ||
            raw.sheets.some(
              (s) => !isRecord(s) || typeof s.name !== "string" || !s.name,
            )
          )
            throw new Error("Workbook must contain named sheets");
          incoming = new FormulaEngine();
          incoming.addWorkbook({
            workbookName:
              typeof raw.name === "string"
                ? raw.name
                : file.name.replace(/\.json$/i, ""),
            data: raw as unknown as WorkbookData,
          });
          incoming = decodeEngine(incoming.serializeEngine());
        } else
          incoming = decodeEngine(
            typeof raw === "string" ? raw : JSON.stringify(raw),
          );
      } else {
        if (!/\.(xlsx|xlsm|csv|tsv|txt)$/i.test(file.name))
          throw new Error(
            "Unsupported file type. Use .xlsx, .csv, .tsv or .json",
          );
        const converted = await spreadsheetToFormulaEngineDetailed(file, {
          fileName: file.name,
        });
        warnings.push(
          ...converted.diagnostics.map((d) => `${file.name}: ${d.message}`),
        );
        incoming = new FormulaEngine();
        incoming.addWorkbook({
          workbookName:
            file.name.replace(/\.[^.]+$/, "") || "Imported workbook",
          data: converted.data,
        });
      }
      const originalNames = [...incoming.getWorkbooks().keys()];
      if (!originalNames.length) throw new Error("File contains no workbooks");
      // Reserve every source name before renaming, so a collision cannot hijack
      // another imported book's references (e.g. A and A (2)).
      const reserved = new Set(
        [...engine.getWorkbooks().keys(), ...originalNames].map((n) =>
          n.toLowerCase(),
        ),
      );
      const names: string[] = [];
      for (const oldName of originalNames) {
        let name = oldName;
        if (
          [...engine.getWorkbooks().keys()].some(
            (n) => n.toLowerCase() === name.toLowerCase(),
          )
        ) {
          for (let i = 2; reserved.has(name.toLowerCase()); i++)
            name = `${oldName} (${i})`;
          incoming.renameWorkbook({
            workbookName: oldName,
            newWorkbookName: name,
          });
          warnings.push(
            `${file.name}: renamed “${oldName}” to “${name}” to keep your existing workbook.`,
          );
        }
        reserved.add(name.toLowerCase());
        names.push(name);
      }
      const merged = mergeEngines(engine, incoming, warnings);
      engine.resetToSerializedEngine(merged);
      result.names.push(...names);
      result.warnings.push(...warnings);
    } catch (error) {
      result.errors.push(`${file.name}: ${message(error)}`);
    }
  }
  return result;
}

export function panes(layout: Layout | null): Pane[] {
  return !layout
    ? []
    : layout.kind === "pane"
      ? [layout.pane]
      : [...panes(layout.first), ...panes(layout.second)];
}
/** Reconcile views with mutable engine data before any pane or formula-bar reads. */
export function reconcileLayout(
  layout: Layout | null,
  engine: FormulaEngine,
): Layout | null {
  if (!layout) return null;
  if (layout.kind === "pane") {
    const { pane } = layout;
    const workbook = engine.getWorkbooks().get(pane.workbookName);
    if (!workbook) return null;
    if (workbook.sheets.has(pane.sheetName)) return layout;
    const sheetName = engine.getOrderedSheetNames(pane.workbookName)[0];
    return sheetName
      ? { ...layout, pane: { ...pane, sheetName, cell: "A1" } }
      : null;
  }
  const first = reconcileLayout(layout.first, engine);
  const second = reconcileLayout(layout.second, engine);
  if (!first || !second) return first ?? second;
  return first === layout.first && second === layout.second
    ? layout
    : { ...layout, first, second };
}
export function splitPane(
  layout: Layout | null,
  targetId: string,
  newPane: Pane,
  direction: "horizontal" | "vertical",
): Layout {
  if (!layout) return { kind: "pane", pane: newPane };
  if (panes(layout).length >= 8)
    throw new Error("A workspace supports at most 8 panes");
  if (panes(layout).some((p) => p.id === newPane.id))
    throw new Error("Pane IDs must be unique");
  const visit = (node: Layout): Layout =>
    node.kind === "pane"
      ? node.pane.id === targetId
        ? {
            kind: "split",
            id: crypto.randomUUID(),
            direction,
            ratio: 0.5,
            first: node,
            second: { kind: "pane", pane: newPane },
          }
        : node
      : { ...node, first: visit(node.first), second: visit(node.second) };
  return visit(layout);
}
export function closePane(layout: Layout | null, id: string): Layout | null {
  if (!layout) return null;
  if (layout.kind === "pane") return layout.pane.id === id ? null : layout;
  const first = closePane(layout.first, id),
    second = closePane(layout.second, id);
  return first && second ? { ...layout, first, second } : (first ?? second);
}
export function updatePane(
  layout: Layout | null,
  id: string,
  patch: Partial<Omit<Pane, "id">>,
): Layout | null {
  if (!layout) return null;
  if (layout.kind === "pane") {
    if (
      layout.pane.id !== id ||
      Object.entries(patch).every(
        ([key, value]) => layout.pane[key as keyof Pane] === value,
      )
    )
      return layout;
    return { ...layout, pane: { ...layout.pane, ...patch } };
  }
  const first = updatePane(layout.first, id, patch)!;
  const second = updatePane(layout.second, id, patch)!;
  return first === layout.first && second === layout.second
    ? layout
    : { ...layout, first, second };
}
export function resizeSplit(
  layout: Layout | null,
  id: string,
  ratio: number,
): Layout | null {
  if (!Number.isFinite(ratio)) throw new Error("Split ratio must be finite");
  if (!layout || layout.kind === "pane") return layout;
  return {
    ...layout,
    ratio:
      layout.id === id ? Math.max(0.15, Math.min(0.85, ratio)) : layout.ratio,
    first: resizeSplit(layout.first, id, ratio)!,
    second: resizeSplit(layout.second, id, ratio)!,
  };
}
function validCell(cell: string): boolean {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/i.exec(cell);
  if (!match) return false;
  const col = [...match[1].toUpperCase()].reduce(
    (n, c) => n * 26 + c.charCodeAt(0) - 64,
    0,
  );
  return col <= 16384 && Number(match[2]) <= 1048576;
}
const isRecord = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === "object" && !Array.isArray(v);
export function restoreSnapshot(value: unknown): LoadedWorkspace {
  if (
    !isRecord(value) ||
    value.format !== "spreadsheet-arena" ||
    value.version !== 1 ||
    typeof value.engine !== "string"
  )
    throw new Error("Unsupported workspace format or version");
  const engine = decodeEngine(value.engine);
  return {
    engine,
    snapshot: {
      format: "spreadsheet-arena",
      version: 1,
      engine: value.engine,
      ...validateWorkspaceView(value, engine),
    },
  };
}
export function validateSnapshot(value: unknown): WorkspaceSnapshot {
  return restoreSnapshot(value).snapshot;
}

/** Validate/detach only the small UI state against the already-live engine. */
export function validateWorkspaceView(
  value: unknown,
  engine: FormulaEngine,
): WorkspaceView {
  if (!isRecord(value)) throw new Error("Invalid workspace view");
  const ids = new Set<string>();
  let count = 0;
  function layout(v: unknown, depth = 0): Layout {
    if (!isRecord(v) || depth > 15) throw new Error("Invalid workspace layout");
    const id = v.kind === "pane" && isRecord(v.pane) ? v.pane.id : v.id;
    if (typeof id !== "string" || !id || ids.has(id))
      throw new Error("Layout IDs must be nonempty and unique");
    ids.add(id);
    if (v.kind === "pane" && isRecord(v.pane)) {
      const p = v.pane;
      if (++count > 8) throw new Error("A workspace supports at most 8 panes");
      if (
        typeof p.workbookName !== "string" ||
        typeof p.sheetName !== "string" ||
        !engine.getWorkbooks().get(p.workbookName)?.sheets.has(p.sheetName)
      )
        throw new Error("Pane refers to a missing workbook or sheet");
      if (
        p.cell !== undefined &&
        (typeof p.cell !== "string" || !validCell(p.cell))
      )
        throw new Error("Invalid selected cell");
      if (
        p.zoom !== undefined &&
        (typeof p.zoom !== "number" ||
          !Number.isFinite(p.zoom) ||
          p.zoom < 0.25 ||
          p.zoom > 4)
      )
        throw new Error("Invalid pane zoom");
      return {
        kind: "pane",
        pane: {
          id,
          workbookName: p.workbookName,
          sheetName: p.sheetName,
          ...(p.cell !== undefined ? { cell: p.cell as string } : {}),
          ...(p.zoom !== undefined ? { zoom: p.zoom as number } : {}),
        },
      };
    }
    if (
      v.kind !== "split" ||
      (v.direction !== "horizontal" && v.direction !== "vertical") ||
      typeof v.ratio !== "number" ||
      !Number.isFinite(v.ratio) ||
      v.ratio < 0.15 ||
      v.ratio > 0.85
    )
      throw new Error("Invalid split layout");
    return {
      kind: "split",
      id,
      direction: v.direction,
      ratio: v.ratio,
      first: layout(v.first, depth + 1),
      second: layout(v.second, depth + 1),
    };
  }
  const tree = value.layout === null ? null : layout(value.layout);
  if (
    value.activePaneId !== null &&
    (typeof value.activePaneId !== "string" ||
      !panes(tree).some((p) => p.id === value.activePaneId))
  )
    throw new Error("Active pane does not exist");
  let view: WorkspaceSnapshot["view"];
  const settings = value.view;
  if (settings !== undefined) {
    if (
      !isRecord(settings) ||
      typeof settings.sidebar !== "boolean" ||
      typeof settings.showFormulas !== "boolean" ||
      (settings.maximized !== null &&
        (typeof settings.maximized !== "string" ||
          !panes(tree).some((p) => p.id === settings.maximized)))
    )
      throw new Error("Invalid workspace view settings");
    view = {
      sidebar: settings.sidebar,
      showFormulas: settings.showFormulas,
      maximized: settings.maximized as string | null,
    };
  }
  return {
    layout: tree,
    activePaneId: value.activePaneId as string | null,
    ...(view ? { view } : {}),
  };
}
export function makeSnapshot(
  engine: FormulaEngine,
  layout: Layout | null,
  activePaneId: string | null,
): WorkspaceSnapshot {
  return {
    format: "spreadsheet-arena",
    version: 1,
    // Workspace backups retain warm evaluations for fast reopening; compact
    // serialization remains useful for standalone workbook interchange.
    engine: engine.serializeEngine(),
    ...validateWorkspaceView({ layout, activePaneId }, engine),
  };
}
