import { describe, expect, it, vi } from "vitest";
import { FormulaEngine, type RangeAddress } from "@ricsam/formula-engine";
import {
  WorkspaceClipboardManager,
  selectionForRanges,
} from "./workspace-clipboard";
const range = (
  workbookName: string,
  row = 0,
  col = 0,
  endRow = row,
  endCol = col,
): RangeAddress => ({
  workbookName,
  sheetName: "Sheet1",
  range: {
    start: { row, col },
    end: {
      row: { type: "number", value: endRow },
      col: { type: "number", value: endCol },
    },
  },
});
function setup() {
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: "Source",
    data: { sheets: [{ name: "Sheet1", content: { A1: 3, B1: "=A1*2" } }] },
  });
  engine.addWorkbook({
    workbookName: "Target",
    data: { sheets: [{ name: "Sheet1", content: { C3: 99 } }] },
  });
  engine.addCellStyle({
    areas: [range("Source", 0, 0, 0, 2)],
    style: { bold: true },
  });
  engine.setCellMetadata(
    { workbookName: "Source", sheetName: "Sheet1", rowIndex: 0, colIndex: 0 },
    { numberFormat: "0.00" },
  );
  engine.addCellDataType({ areas: [range("Source")], dataType: "number" });
  const clipboard = new WorkspaceClipboardManager(engine);
  return { engine, clipboard };
}
describe("workspace clipboard", () => {
  it("copies only the rows a table filter shows and pastes them without gaps", () => {
    const engine = new FormulaEngine();
    engine.addWorkbook({
      workbookName: "Book",
      data: {
        sheets: [
          {
            name: "Sheet1",
            content: {
              A1: "Item",
              B1: "Cost",
              A2: "a",
              B2: 1,
              A3: "b",
              B3: 2,
              A4: "c",
              B4: 3,
              A5: "d",
              B5: "=B4+1",
            },
          },
        ],
        tables: [
          {
            name: "Costs",
            sheetName: "Sheet1",
            start: "A1",
            numRows: { type: "number", value: 4 },
            numCols: 2,
          },
        ],
      },
    });
    engine.setTableColumnFilter({
      workbookName: "Book",
      tableName: "Costs",
      column: 0,
      filter: { type: "values", values: ["a", "c", "d"] },
    });
    const clipboard = new WorkspaceClipboardManager(engine);
    clipboard.copyRanges([range("Book", 1, 0, 4, 1)]);
    clipboard.pasteRanges([range("Book", 0, 3)]);
    const sheet = { workbookName: "Book", sheetName: "Sheet1" };
    const value = (rowIndex: number, colIndex: number) =>
      engine.getCellValue({ ...sheet, rowIndex, colIndex });
    expect([0, 1, 2].map((row) => [value(row, 3), value(row, 4)])).toEqual([
      ["a", 1],
      ["c", 3],
      ["d", 4],
    ]);
    // The formula moved up with its row, like the values next to it.
    const content = engine.getSheet(sheet)!.content;
    expect(content.get("E3")).toBe("=E2+1");
    expect(content.has("D4") || content.has("E4")).toBe(false);
    engine.undo();
    expect(engine.getSheet(sheet)!.content.has("D1")).toBe(false);
  });
  it("blocks clipboard mutation callbacks while a save lock is held", () => {
    const { engine } = setup();
    let locked = false;
    const clipboard = new WorkspaceClipboardManager(engine, () => locked);
    clipboard.copyRanges([range("Source")]);
    locked = true;
    clipboard.pasteRanges([range("Target", 2, 2)]);
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("C3"),
    ).toBe(99);
    locked = false;
    clipboard.pasteRanges([range("Target", 2, 2)]);
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("C3"),
    ).toBe(3);
  });
  it("copies a blank styled cell and clears a strongly styled destination with plain source", () => {
    const { engine, clipboard } = setup();
    const dest = {
      workbookName: "Target",
      sheetName: "Sheet1",
      rowIndex: 2,
      colIndex: 2,
    };
    clipboard.copyRanges([range("Source", 0, 2)]);
    clipboard.pasteRanges([range("Target", 2, 2)]);
    expect(engine.getSheet(dest)!.content.get("C3")).toBeUndefined();
    expect(engine.getCellStyle(dest)).toEqual({ bold: true });
    engine.addCellStyle({
      areas: [range("Target", 2, 2)],
      style: {
        italic: true,
        backgroundColor: "#ff0000",
        borderColor: "#000000",
        borderSides: { top: true },
      },
    });
    clipboard.copyRanges([range("Source", 3, 3)]);
    clipboard.pasteRanges([range("Target", 2, 2)]);
    expect(engine.getCellStyle(dest)).toBeUndefined();
    expect(engine.getCellDataType(dest)).toBe("general");
  });
  it.each([
    [4, 4],
    [4, 1],
    [1, 4],
  ])("matches engine finite tiling for a %i by %i target", (h, w) => {
    const { engine, clipboard } = setup();
    engine.setCellContent(
      { workbookName: "Source", sheetName: "Sheet1", rowIndex: 1, colIndex: 0 },
      7,
    );
    engine.setCellContent(
      { workbookName: "Source", sheetName: "Sheet1", rowIndex: 1, colIndex: 1 },
      8,
    );
    clipboard.copyRanges([range("Source", 0, 0, 1, 1)]);
    const expected = new FormulaEngine();
    expected.resetToSerializedEngine(engine.serializeEngine());
    const target = range("Target", 4, 4, 4 + h - 1, 4 + w - 1);
    expected.smartPaste(
      clipboard.copiedCells,
      { ...target, areas: [target.range] },
      { type: "formula", include: "all" },
    );
    clipboard.pasteRanges([target]);
    expect(engine.getSheet(target)!.content).toEqual(
      expected.getSheet(target)!.content,
    );
  });
  it("cuts without formatting, preserving destination and source styles", () => {
    const { engine, clipboard } = setup();
    const dest = {
      workbookName: "Target",
      sheetName: "Sheet1",
      rowIndex: 2,
      colIndex: 2,
    };
    engine.addCellStyle({
      areas: [range("Target", 2, 2)],
      style: { italic: true },
    });
    clipboard.copyRanges([range("Source")], true);
    clipboard.pasteRanges([range("Target", 2, 2)], "without-formatting");
    expect(engine.getCellStyle(dest)).toEqual({ italic: true });
    expect(engine.getCellValue(dest)).toBe(3);
    expect(
      engine.getCellStyle({
        ...dest,
        workbookName: "Source",
        rowIndex: 0,
        colIndex: 0,
      }),
    ).toEqual({ bold: true });
  });
  it("pastes formula snapshot after removing source, including FormulaSheet's formula paste type", () => {
    const { engine, clipboard } = setup();
    clipboard.copyRanges([range("Source", 0, 0, 0, 1)]);
    engine.removeWorkbook("Source");
    clipboard.triggerPaste({
      workbookName: "Target",
      sheetName: "Sheet1",
      selectionManager: selectionForRanges([range("Target", 2, 2)]),
      rawString: clipboard.signature,
      updates: [],
      pasteType: "formula",
    });
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("D3"),
    ).toBe("=C3*2");
  });
  it("rejects overflowing destinations atomically and preserves the buffer", () => {
    const { engine, clipboard } = setup();
    clipboard.copyRanges([range("Source", 0, 0, 0, 1)]);
    engine.clearUndoRedoHistory();
    expect(() => clipboard.pasteRanges([range("Target", 0, 16383)])).toThrow(
      "worksheet limits",
    );
    expect(engine.canUndo()).toBe(false);
    expect(clipboard.hasContent).toBe(true);
    clipboard.pasteRanges([range("Target", 1048575, 16382)]);
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("XFD1048576"),
    ).toBe("=XFC1048576*2");
    expect(() => clipboard.pasteRanges([range("Missing")])).toThrow(
      "Worksheet not found",
    );
  });
  it("rejects oversized finite and open ranges before extracting cells", () => {
    const { engine, clipboard } = setup();
    const extract = vi.spyOn(clipboard, "extractCellsFromSelection");
    expect(() =>
      clipboard.copyRanges([range("Source", 0, 0, 100_000, 1)]),
    ).toThrow("100,000");
    expect(extract).not.toHaveBeenCalled();
    engine.setCellContent(
      {
        workbookName: "Source",
        sheetName: "Sheet1",
        rowIndex: 100_000,
        colIndex: 0,
      },
      1,
    );
    const open = range("Source");
    open.range.end.row = { type: "infinity", sign: "positive" };
    expect(() => clipboard.copyRanges([open])).toThrow("100,000");
    expect(extract).not.toHaveBeenCalled();
  });
  it("snapshots a range including blanks, translates formulas cross-workbook, copies styles and types, undoes once", () => {
    const { engine, clipboard } = setup();
    clipboard.copyRanges([range("Source", 0, 0, 0, 2)]);
    engine.setCellContent(
      { workbookName: "Source", sheetName: "Sheet1", rowIndex: 0, colIndex: 0 },
      42,
    );
    engine.clearCellStyles(range("Source", 0, 0, 0, 2));
    engine.clearUndoRedoHistory();
    clipboard.pasteRanges([range("Target", 2, 2)]);
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("C3"),
    ).toBe(3);
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("D3"),
    ).toBe("=C3*2");
    expect(
      engine.getCellStyle({
        workbookName: "Target",
        sheetName: "Sheet1",
        rowIndex: 2,
        colIndex: 4,
      }),
    ).toMatchObject({ bold: true });
    expect(
      engine.getCellMetadata({
        workbookName: "Target",
        sheetName: "Sheet1",
        rowIndex: 2,
        colIndex: 2,
      }),
    ).toEqual({ numberFormat: "0.00" });
    expect(
      engine.getCellDataType({
        workbookName: "Target",
        sheetName: "Sheet1",
        rowIndex: 2,
        colIndex: 2,
      }),
    ).toBe("number");
    expect(engine.undo()).toBe(true);
    expect(engine.canUndo()).toBe(false);
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("C3"),
    ).toBe(99);
  });
  it("without formatting preserves destination style and data type, retaining formulas", () => {
    const { engine, clipboard } = setup();
    engine.addCellStyle({
      areas: [range("Target", 2, 2, 2, 3)],
      style: { italic: true },
    });
    engine.addCellDataType({
      areas: [range("Target", 2, 2)],
      dataType: "text",
    });
    clipboard.copyRanges([range("Source", 0, 0, 0, 1)]);
    clipboard.pasteRanges([range("Target", 2, 2)], "without-formatting");
    expect(
      engine.getCellStyle({
        workbookName: "Target",
        sheetName: "Sheet1",
        rowIndex: 2,
        colIndex: 2,
      }),
    ).toEqual({ italic: true });
    expect(
      engine.getCellDataType({
        workbookName: "Target",
        sheetName: "Sheet1",
        rowIndex: 2,
        colIndex: 2,
      }),
    ).toBe("text");
    expect(
      engine
        .getSheet({ workbookName: "Target", sheetName: "Sheet1" })!
        .content.get("D3"),
    ).toBe("=C3*2");
  });
  it("retains keyboard internal paste and delegates external text updates", () => {
    const { clipboard } = setup();
    clipboard.copyRanges([range("Source")]);
    const selectionManager = selectionForRanges([range("Target")]);
    const save = vi.spyOn(selectionManager, "saveCellValues");
    const context = {
      workbookName: "Target",
      sheetName: "Sheet1",
      selectionManager,
      updates: [{ rowIndex: 0, colIndex: 0, value: "external" }],
      rawString: "external",
      pasteType: "formula" as const,
    };
    clipboard.triggerPaste(context);
    expect(save).toHaveBeenCalledWith(context.updates);
    clipboard.triggerPaste({ ...context, rawString: clipboard.signature });
    expect(save).toHaveBeenCalledTimes(1);
  });
  it("cuts with a single engine undo and consumes the move buffer", () => {
    const { engine, clipboard } = setup();
    clipboard.copyRanges([range("Source")], true);
    engine.clearUndoRedoHistory();
    clipboard.pasteRanges([range("Target", 1, 1)]);
    expect(
      engine
        .getSheet({ workbookName: "Source", sheetName: "Sheet1" })!
        .content.has("A1"),
    ).toBe(false);
    expect(clipboard.hasContent).toBe(false);
    expect(engine.undo()).toBe(true);
    expect(engine.canUndo()).toBe(false);
    expect(
      engine
        .getSheet({ workbookName: "Source", sheetName: "Sheet1" })!
        .content.get("A1"),
    ).toBe(3);
  });
});
