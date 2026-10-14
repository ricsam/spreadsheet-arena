import { describe, expect, it } from "vitest";
import { FormulaEngine } from "@ricsam/formula-engine";
import { arenaTools } from "./webmcp";
import { createEmptyWorkbook } from "./model";
import { cellAddress, parseInput } from "./cells";

function setup() {
  const engine = FormulaEngine.buildEmpty();
  createEmptyWorkbook(engine, "Book");
  const tools = arenaTools({
    engine,
    layout: null,
    activePaneId: null,
    createWorkbook: (name) => createEmptyWorkbook(engine, name),
    focusCell: (workbook, sheet, cell) => ({ workbook, sheet, cell }),
  });
  const call = (name: string, args: Record<string, unknown>) =>
    tools.find((t) => t.name === name)!.execute(args) as Promise<{
      isError?: boolean;
      content: { text: string }[];
    }>;
  return { engine, tools, call };
}
describe("automation contracts", () => {
  it("lists workbooks and edits typed cell values with errors surfaced", async () => {
    const { call } = setup();
    expect(
      JSON.parse((await call("list_workbooks", {})).content[0].text).workbooks,
    ).toEqual([{ name: "Book", sheets: ["Sheet1"] }]);
    await call("set_cell", {
      workbook: "Book",
      sheet: "Sheet1",
      cell: "A1",
      value: 7,
    });
    await call("set_cell", {
      workbook: "Book",
      sheet: "Sheet1",
      cell: "B1",
      value: "=A1*3",
    });
    const result = await call("read_cells", {
      workbook: "Book",
      sheet: "Sheet1",
      cells: ["A1", "B1"],
    });
    expect(
      JSON.parse(result.content[0].text).map(
        (cell: { value: number }) => cell.value,
      ),
    ).toEqual([7, 21]);
    await call("set_cell", {
      workbook: "Book",
      sheet: "Sheet1",
      cell: "A1",
      value: null,
    });
    expect(
      (
        await call("set_cell", {
          workbook: "Book",
          sheet: "missing",
          cell: "A1",
          value: 1,
        })
      ).isError,
    ).toBe(true);
    expect(
      (
        await call("set_cell", {
          workbook: "Book",
          sheet: "Sheet1",
          cell: "A1",
          value: {},
        })
      ).isError,
    ).toBe(true);
  });
  it("bounds reads and validates coordinates before changing selection", async () => {
    const { call } = setup();
    expect(
      (
        await call("read_cells", {
          workbook: "Book",
          sheet: "Sheet1",
          cells: Array(201).fill("A1"),
        })
      ).isError,
    ).toBe(true);
    expect(
      (
        await call("select_cell", {
          workbook: "Book",
          sheet: "Sheet1",
          cell: "A0",
        })
      ).isError,
    ).toBe(true);
    expect(
      (
        await call("select_cell", {
          workbook: "missing",
          sheet: "Sheet1",
          cell: "A1",
        })
      ).isError,
    ).toBe(true);
    expect(
      (await call("create_workbook", { name: "x".repeat(101) })).isError,
    ).toBe(true);
    expect(
      JSON.parse(
        (await call("create_workbook", { name: "New" })).content[0].text,
      ).workbook,
    ).toBe("New");
  });
  it("marks only nonmutating tools read-only", () => {
    expect(
      setup()
        .tools.filter((t) => t.annotations?.readOnlyHint)
        .map((t) => t.name),
    ).toEqual(["list_workbooks", "read_cells"]);
  });
});
describe("formula bar parsing", () => {
  it("preserves leading zeros and formulas and supports numbers, booleans and clearing", () => {
    expect(parseInput("0012")).toBe("0012");
    expect(parseInput("=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
    expect(parseInput("42.5")).toBe(42.5);
    expect(parseInput("TRUE")).toBe(true);
    expect(parseInput("")).toBeUndefined();
  });
  it("accepts Excel boundaries and rejects out-of-range addresses", () => {
    expect(cellAddress("B", "S", "xfd1048576")).toMatchObject({
      rowIndex: 1048575,
      colIndex: 16383,
    });
    expect(() => cellAddress("B", "S", "XFE1")).toThrow();
    expect(() => cellAddress("B", "S", "A0")).toThrow();
  });
});
