import { describe, expect, it } from "vitest";
import { FormulaEngine } from "@ricsam/formula-engine";
import {
  MAX_IMPORT_FILES,
  MAX_IMPORT_FILE_BYTES,
  closePane,
  createDemoEngine,
  createEmptyWorkbook,
  decodeEngine,
  exportWorkbook,
  importFiles,
  makeSnapshot,
  panes,
  resizeSplit,
  splitPane,
  uniqueWorkbookName,
  updatePane,
  validateSnapshot,
  type Layout,
} from "./model";

const pane: Layout = {
  kind: "pane",
  pane: {
    id: "one",
    workbookName: "Launch budget",
    sheetName: "Overview",
    cell: "D12",
    zoom: 1.2,
  },
};
const jsonFile = (value: unknown, name = "books.json") =>
  new File([typeof value === "string" ? value : JSON.stringify(value)], name, {
    type: "application/json",
  });
const cell = (
  workbookName: string,
  sheetName: string,
  colIndex: number,
  rowIndex: number,
) => ({ workbookName, sheetName, colIndex, rowIndex });

describe("real formula engine state", () => {
  it("demo evaluates cross-workbook references and responds to input edits", () => {
    const e = createDemoEngine();
    expect(
      e.getCellValue(cell("Launch budget", "Overview", 3, 11)),
    ).toBeCloseTo(7535);
    e.setCellContent(cell("Assumptions", "Inputs", 1, 3), 50);
    expect(
      e.getCellValue(cell("Launch budget", "Overview", 3, 11)),
    ).toBeCloseTo(7975);
    expect(
      decodeEngine(e.serializeEngine()).getCellValue(
        cell("Launch budget", "Overview", 3, 11),
      ),
    ).toBeCloseTo(7975);
  });
  it("creates unique empty workbook names without replacing existing books", () => {
    const e = createDemoEngine();
    expect(uniqueWorkbookName(e, "launch budget")).toBe("launch budget (2)");
    expect(createEmptyWorkbook(e, "Launch budget")).toBe("Launch budget (2)");
    expect(
      e.getWorkbooks().get("Launch budget (2)")?.sheets.has("Sheet1"),
    ).toBe(true);
  });
  it("exports one native workbook without altering other workbooks or cross-book formulas", () => {
    const e = createDemoEngine();
    const before = e.serializeEngine();
    e.setCellMetadata(cell("Launch budget", "Overview", 0, 0), {
      note: "Keep this",
    });
    const exported = decodeEngine(exportWorkbook(e, "Launch budget"));
    expect([...exported.getWorkbooks().keys()]).toEqual(["Launch budget"]);
    expect(
      exported.getCellMetadata(cell("Launch budget", "Overview", 0, 0)),
    ).toEqual({ note: "Keep this" });
    expect(
      exported.getCellStyle(cell("Launch budget", "Overview", 0, 0))?.bold,
    ).toBe(true);
    expect(e.getWorkbooks().size).toBe(2);
    expect(e.serializeEngine()).not.toBe(before); // metadata was deliberately edited
    expect(
      exported
        .getWorkbooks()
        .get("Launch budget")
        ?.sheets.get("Overview")
        ?.content.get("C5"),
    ).toBe("=[Assumptions]Inputs!B4");
  });
  it("rejects malformed native snapshots on a disposable engine", () => {
    for (const data of [
      "{}",
      "null",
      "{",
      '{"version":6,"managers":{"style":{"cellDataTypes":[]}}}',
    ])
      expect(() => decodeEngine(data)).toThrow(
        "Cannot read spreadsheet engine",
      );
  });
  it("imports colliding engine books atomically and keeps imported relationships", async () => {
    const e = createDemoEngine(),
      source = createDemoEngine();
    source.setCellContent(cell("Assumptions", "Inputs", 1, 3), 100);
    const result = await importFiles(e, [jsonFile(source.serializeEngine())]);
    expect(result.errors).toEqual([]);
    expect(result.names).toEqual(["Assumptions (2)", "Launch budget (2)"]);
    expect(
      e.getCellValue(cell("Launch budget", "Overview", 3, 11)),
    ).toBeCloseTo(7535);
    expect(
      e.getCellValue(cell("Launch budget (2)", "Overview", 3, 11)),
    ).toBeCloseTo(12375);
    e.setCellContent(cell("Assumptions (2)", "Inputs", 1, 3), 50);
    expect(
      e.getCellValue(cell("Launch budget (2)", "Overview", 3, 11)),
    ).toBeCloseTo(7975);
    expect(
      e.getCellValue(cell("Launch budget", "Overview", 3, 11)),
    ).toBeCloseTo(7535);
    expect(
      e
        .getWorkbooks()
        .get("Launch budget (2)")
        ?.sheets.get("Overview")
        ?.content.get("C5"),
    ).toBe("=[Assumptions (2)]Inputs!B4");
    expect(result.warnings).toHaveLength(2);
  });
  it("preserves tables, metadata and styles through export/import", async () => {
    const e = new FormulaEngine();
    e.addWorkbook({
      workbookName: "Products",
      data: {
        workbookMetadata: { source: "test" },
        sheets: [
          {
            name: "Data",
            content: { A1: "Name", B1: "Cost", A2: "Tea", B2: 4 },
            sheetMetadata: { width: 100 },
          },
        ],
        tables: [
          {
            name: "ProductsTable",
            sheetName: "Data",
            start: "A1",
            numRows: { type: "number", value: 1 },
            numCols: 2,
          },
        ],
      },
    });
    const target = new FormulaEngine();
    const result = await importFiles(target, [
      jsonFile(exportWorkbook(e, "Products")),
    ]);
    expect(result.errors).toEqual([]);
    expect(target.getWorkbookMetadata("Products")).toEqual({ source: "test" });
    expect(target.getState().tables.get("Products")?.has("ProductsTable")).toBe(
      true,
    );
  });
  it("does not corrupt current state on failed import, and continues later files", async () => {
    const e = createDemoEngine();
    const before = e.serializeEngine();
    const bad = await importFiles(e, [jsonFile("{}")]);
    expect(bad.errors).toHaveLength(1);
    expect(e.serializeEngine()).toBe(before);
    const result = await importFiles(e, [
      jsonFile("{"),
      new File(["Name,Amount\nTea,12\nCoffee,20"], "Sales.csv"),
    ]);
    expect(result.errors).toHaveLength(1);
    expect(result.names).toEqual(["Sales"]);
  });
  it("imports a real Excel workbook through the converter", async () => {
    const { default: ExcelJS } = await import("exceljs");
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Forecast");
    sheet.addRow(["Month", "Revenue"]);
    sheet.addRow(["January", 1200]);
    const bytes = await workbook.xlsx.writeBuffer();
    const engine = new FormulaEngine();
    const result = await importFiles(engine, [
      new File([new Uint8Array(bytes)], "Forecast.xlsx"),
    ]);
    expect(result.errors).toEqual([]);
    expect(result.names).toEqual(["Forecast"]);
    expect(engine.getCellValue(cell("Forecast", "Forecast", 1, 1))).toBe(1200);
  });
  it("imports standalone workbook data and workspace JSON as books only", async () => {
    const e = new FormulaEngine();
    expect(
      (
        await importFiles(e, [
          jsonFile(
            { sheets: [{ name: "Sheet1", content: { A1: 42 } }] },
            "Answer.json",
          ),
        ])
      ).names,
    ).toEqual(["Answer"]);
    const snapshot = makeSnapshot(createDemoEngine(), pane, "one");
    expect((await importFiles(e, [jsonFile(snapshot)])).errors).toEqual([]);
    expect(e.getWorkbooks().size).toBe(3);
  });
  it("rejects oversized batches/files and corrupted Excel without touching live state", async () => {
    const e = createDemoEngine(),
      before = e.serializeEngine();
    const batch = await importFiles(
      e,
      Array.from(
        { length: MAX_IMPORT_FILES + 1 },
        () => new File(["a,b"], "x.csv"),
      ),
    );
    expect(batch.errors[0]).toContain("at most");
    const large = new File(["x"], "large.csv");
    Object.defineProperty(large, "size", { value: MAX_IMPORT_FILE_BYTES + 1 });
    expect((await importFiles(e, [large])).errors[0]).toContain("25 MB");
    expect(
      (await importFiles(e, [new File(["not an Excel file"], "broken.xlsx")]))
        .errors[0],
    ).toContain("not one");
    expect(e.serializeEngine()).toBe(before);
  });
  it("warns on global name collisions and retains existing definitions", async () => {
    const e = createDemoEngine(),
      source = new FormulaEngine();
    createEmptyWorkbook(source, "New");
    e.addNamedExpression({ expressionName: "TaxRate", expression: "0.2" });
    source.addNamedExpression({ expressionName: "TaxRate", expression: "0.5" });
    const result = await importFiles(e, [jsonFile(source.serializeEngine())]);
    expect(result.errors).toEqual([]);
    expect(
      result.warnings.some((w) => w.includes("TaxRate") && w.includes("kept")),
    ).toBe(true);
    expect(
      e.getState().namedExpressions.globalExpressions.get("TaxRate")
        ?.expression,
    ).toBe("0.2");
  });
  it("exports only reachable global names, including transitive dependencies", () => {
    const e = createDemoEngine();
    e.addNamedExpression({ expressionName: "BaseRate", expression: "2" });
    e.addNamedExpression({
      expressionName: "UsedRate",
      expression: "BaseRate*3",
    });
    e.addNamedExpression({
      expressionName: "PrivateUnrelated",
      expression: "999",
    });
    e.setCellContent(cell("Launch budget", "Overview", 4, 0), "=UsedRate");
    const exported = decodeEngine(exportWorkbook(e, "Launch budget"));
    expect(
      [...exported.getState().namedExpressions.globalExpressions.keys()].sort(),
    ).toEqual(["BaseRate", "UsedRate"]);
    expect(exported.getCellValue(cell("Launch budget", "Overview", 4, 0))).toBe(
      6,
    );
  });
  it("filters exported references and safely remaps imported reference IDs", async () => {
    const e = createDemoEngine();
    const raw = JSON.parse(e.serializeEngine());
    const ref = (id: string, workbookName: string, sheetName: string) => [
      id,
      {
        id,
        isValid: true,
        address: {
          workbookName,
          sheetName,
          range: {
            start: { row: 0, col: 0 },
            end: {
              row: { type: "number", value: 0 },
              col: { type: "number", value: 0 },
            },
          },
        },
      },
    ];
    raw.managers.reference.entries = [
      ref("keep", "Launch budget", "Overview"),
      ref("private", "Assumptions", "Inputs"),
    ];
    const source = decodeEngine(JSON.stringify(raw));
    const exported = JSON.parse(exportWorkbook(source, "Launch budget"));
    expect(
      exported.managers.reference.entries.map(([id]: [string, unknown]) => id),
    ).toEqual(["keep"]);
    const result = await importFiles(source, [
      jsonFile(source.serializeEngine()),
    ]);
    expect(result.errors).toEqual([]);
    const refs = source.getState().references;
    expect(refs.size).toBe(4);
    expect(refs.get("keep")?.address.workbookName).toBe("Launch budget");
    expect(
      [...refs.values()].some(
        (r) => r.address.workbookName === "Launch budget (2)",
      ),
    ).toBe(true);
  });
  it("reserves imported names before collision renaming", async () => {
    const e = new FormulaEngine(),
      incoming = new FormulaEngine();
    createEmptyWorkbook(e, "A");
    createEmptyWorkbook(incoming, "A");
    createEmptyWorkbook(incoming, "A (2)");
    const result = await importFiles(e, [jsonFile(incoming.serializeEngine())]);
    expect(result.errors).toEqual([]);
    expect(result.names).toEqual(["A (3)", "A (2)"]);
  });
});

describe("workspace layout", () => {
  it("round-trips selection, zoom, engine and nested panes", () => {
    const tree = splitPane(
      pane,
      "one",
      { id: "two", workbookName: "Assumptions", sheetName: "Inputs" },
      "vertical",
    );
    const s = makeSnapshot(createDemoEngine(), tree, "two");
    expect(validateSnapshot(JSON.parse(JSON.stringify(s)))).toEqual(s);
    expect(panes(s.layout)[0].cell).toBe("D12");
    expect(panes(s.layout)[0].zoom).toBe(1.2);
  });
  it("updates immutably, clamps resizing and collapses closed branches", () => {
    const tree = splitPane(
      pane,
      "one",
      { id: "two", workbookName: "Assumptions", sheetName: "Inputs" },
      "horizontal",
    );
    expect(panes(updatePane(tree, "two", { cell: "A4" }))[1].cell).toBe("A4");
    expect(panes(tree)[1].cell).toBeUndefined();
    if (tree.kind !== "split") throw new Error("expected split");
    expect(resizeSplit(tree, tree.id, 2)).toMatchObject({ ratio: 0.85 });
    expect(() => resizeSplit(tree, tree.id, NaN)).toThrow();
    expect(closePane(tree, "two")).toEqual(pane);
    expect(closePane(closePane(tree, "two"), "one")).toBeNull();
  });
  it("rejects missing sheets, malformed selection, active IDs, ratios and duplicate IDs", () => {
    const s = makeSnapshot(createDemoEngine(), pane, "one");
    expect(() => validateSnapshot({ ...s, activePaneId: "missing" })).toThrow();
    expect(() =>
      validateSnapshot({
        ...s,
        layout: updatePane(pane, "one", { sheetName: "missing" }),
      }),
    ).toThrow();
    expect(() =>
      validateSnapshot({
        ...s,
        layout: updatePane(pane, "one", { cell: "A0" }),
      }),
    ).toThrow();
    expect(() =>
      validateSnapshot({
        ...s,
        layout: updatePane(pane, "one", { zoom: Infinity }),
      }),
    ).toThrow();
    expect(() =>
      validateSnapshot({
        ...s,
        layout: {
          kind: "split",
          id: "s",
          direction: "horizontal",
          ratio: 0.5,
          first: pane,
          second: pane,
        },
      }),
    ).toThrow();
  });
  it("bounds selected cells to the Excel grid and split ratios to UI limits", () => {
    const s = makeSnapshot(createDemoEngine(), pane, "one");
    for (const selected of ["XFE1", "A1048577", "ZZZZ1", "A99999999999999"]) {
      expect(() =>
        validateSnapshot({
          ...s,
          layout: updatePane(pane, "one", { cell: selected }),
        }),
      ).toThrow("selected cell");
    }
    expect(() =>
      validateSnapshot({
        ...s,
        layout: updatePane(pane, "one", { cell: "XFD1048576" }),
      }),
    ).not.toThrow();
    const tree = splitPane(
      pane,
      "one",
      { id: "two", workbookName: "Assumptions", sheetName: "Inputs" },
      "horizontal",
    );
    expect(() =>
      validateSnapshot({ ...s, layout: { ...tree, ratio: 0.1 } }),
    ).toThrow("split");
    if (tree.kind === "split")
      expect(resizeSplit(tree, tree.id, -1)).toMatchObject({ ratio: 0.15 });
  });
  it("caps panes at eight", () => {
    let tree: Layout = pane;
    for (let i = 2; i <= 8; i++)
      tree = splitPane(
        tree,
        "one",
        { id: String(i), workbookName: "Launch budget", sheetName: "Overview" },
        "horizontal",
      );
    expect(panes(tree)).toHaveLength(8);
    expect(() =>
      splitPane(
        tree,
        "one",
        { id: "9", workbookName: "Launch budget", sheetName: "Overview" },
        "vertical",
      ),
    ).toThrow("at most 8");
    const nine = {
      kind: "split",
      id: "extra",
      direction: "vertical",
      ratio: 0.5,
      first: tree,
      second: {
        kind: "pane",
        pane: { id: "9", workbookName: "Launch budget", sheetName: "Overview" },
      },
    };
    expect(() =>
      validateSnapshot({
        ...makeSnapshot(createDemoEngine(), null, null),
        layout: nine,
      }),
    ).toThrow("at most 8");
  });
});
