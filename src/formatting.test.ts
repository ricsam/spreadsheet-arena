import { describe, expect, it } from "vitest";
import { FormulaEngine, type RangeAddress } from "@ricsam/formula-engine";
import {
  applyStyle,
  applyDataType,
  saveTable,
  setDimension,
  tableRange,
  validateTableName,
} from "./formatting";
const location = { workbookName: "Book", sheetName: "Sheet1" };
const range = (r = 2, c = 1): RangeAddress => ({
  ...location,
  range: {
    start: { row: 0, col: 0 },
    end: {
      row: { type: "number", value: r },
      col: { type: "number", value: c },
    },
  },
});
function engine() {
  const e = new FormulaEngine();
  e.addWorkbook({
    workbookName: "Book",
    data: {
      sheets: [
        {
          name: "Sheet1",
          content: { A1: "Name", B1: "Value", A2: "One", B2: 10, B3: 20 },
          sheetMetadata: { keep: true, arenaColumnWidths: { Z: 130 } },
        },
      ],
    },
  });
  e.clearUndoRedoHistory();
  return e;
}
describe("range formatting", () => {
  it("layers only changed properties and undoes in one entry", () => {
    const e = engine(),
      a = { ...location, rowIndex: 0, colIndex: 0 };
    applyStyle(e, [range()], { color: "#123456", italic: true });
    applyStyle(e, [range()], { bold: true });
    expect(e.getCellStyle(a)).toMatchObject({
      color: "#123456",
      italic: true,
      bold: true,
    });
    e.undo();
    expect(e.getCellStyle(a)).toEqual({ color: "#123456", italic: true });
    e.redo();
    expect(e.getCellStyle(a)?.bold).toBe(true);
  });
  it("formats whole columns sparsely and preserves values", () => {
    const e = engine(),
      area = range();
    area.range.end.row = { type: "infinity", sign: "positive" };
    applyStyle(e, [area], { wrapText: true });
    applyDataType(e, [area], "number");
    expect(e.getAllCellStyles()).toHaveLength(1);
    expect(e.getSheet(location)?.content.size).toBe(5);
    expect(
      e.getCellDataType({ ...location, rowIndex: 100000, colIndex: 0 }),
    ).toBe("number");
  });
  it("preserves sheet metadata and undo for dimensions, rejects unbounded dimensions", () => {
    const e = engine();
    setDimension(e, [range()], "col", 170);
    expect(e.getSheetMetadata(location)).toEqual({
      keep: true,
      arenaColumnWidths: { A: 170, B: 170, Z: 130 },
    });
    e.undo();
    expect(e.getSheetMetadata(location)).toEqual({
      keep: true,
      arenaColumnWidths: { Z: 130 },
    });
    const area = range();
    area.range.end.row = { type: "infinity", sign: "positive" };
    expect(() => setDimension(e, [area], "row", 40)).toThrow();
    expect(() => applyStyle(e, [range()], { fontSize: NaN })).toThrow();
  });
});
describe("table definitions", () => {
  it("creates exact finite geometry, resizes, renames, removes without removing values", () => {
    const e = engine();
    saveTable(e, [range()], "Items");
    const table = e.getTable({ workbookName: "Book", tableName: "Items" })!;
    expect(tableRange(table)).toEqual(range());
    saveTable(e, [range(1)], "Renamed", table);
    expect(
      e.getTable({ workbookName: "Book", tableName: "Renamed" })?.endRow,
    ).toEqual({ type: "number", value: 1 });
    e.undo();
    expect(
      e.getTable({ workbookName: "Book", tableName: "Items" })?.endRow,
    ).toEqual({ type: "number", value: 2 });
    const content = [...e.getSheet(location)!.content];
    e.removeTable({ workbookName: "Book", tableName: "Items" });
    expect([...e.getSheet(location)!.content]).toEqual(content);
    e.undo();
    expect(e.hasTable({ workbookName: "Book", tableName: "Items" })).toBe(true);
  });
  it("rejects overlaps, invalid names, duplicates, nonrectangular and unbounded selections", () => {
    const e = engine();
    saveTable(e, [range()], "Items");
    expect(() => saveTable(e, [range()], "Other")).toThrow(/overlap/);
    expect(() => validateTableName(e, "Book", "items")).toThrow();
    expect(() => validateTableName(e, "Book", "A1")).toThrow();
    expect(() => saveTable(e, [range(), range()], "Other")).toThrow();
    const area = range();
    area.range.end.row = { type: "infinity", sign: "positive" };
    expect(() => saveTable(e, [area], "Other")).toThrow();
    expect(e.getTables("Book").size).toBe(1);
  });
});
