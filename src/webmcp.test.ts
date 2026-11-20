import { describe, expect, it } from "vitest";
import { FormulaEngine } from "@ricsam/formula-engine";
import { arenaTools } from "./webmcp";
import { createEmptyWorkbook } from "./model";
import { cellAddress, parseInput } from "./cells";

function setup(isLocked?: () => boolean) {
  const engine = FormulaEngine.buildEmpty();
  createEmptyWorkbook(engine, "Book");
  const tools = arenaTools({
    engine,
    isLocked,
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
  it("checks the live save lock even for previously registered tools", async () => {
    let locked = true;
    const { engine, call } = setup(() => locked);
    const input = { workbook: "Book", sheet: "Sheet1", cell: "A1", value: 9 };
    expect((await call("set_cell", input)).isError).toBe(true);
    expect((await call("create_workbook", { name: "Blocked" })).isError).toBe(
      true,
    );
    expect((await call("select_cell", input)).isError).toBe(true);
    expect(
      (await call("read_cells", { ...input, cells: ["A1"] })).isError,
    ).toBe(true);
    expect(engine.getWorkbooks().size).toBe(1);
    locked = false;
    expect((await call("set_cell", input)).isError).toBeUndefined();
    expect(engine.getCellValue(cellAddress("Book", "Sheet1", "A1"))).toBe(9);
  });
  it("marks only nonmutating tools read-only", () => {
    expect(
      setup()
        .tools.filter((t) => t.annotations?.readOnlyHint)
        .map((t) => t.name),
    ).toEqual(["list_workbooks", "read_cells", "list_tables"]);
  });
  it("sorts and filters tables and reports the rows filters hide", async () => {
    const { engine, call } = setup();
    const sheet = { workbookName: "Book", sheetName: "Sheet1" };
    engine.setSheetContent(
      sheet,
      new Map<string, string | number>([
        ["A1", "Item"],
        ["B1", "Amount"],
        ["A2", "b"],
        ["B2", 2],
        ["A3", "a"],
        ["B3", 1],
        ["A4", "c"],
        ["B4", 3],
      ]),
    );
    engine.addTable({
      ...sheet,
      tableName: "Costs",
      start: "A1",
      numCols: 2,
      numRows: { type: "number", value: 3 },
    });
    const json = async (name: string, args: Record<string, unknown>) =>
      JSON.parse((await call(name, args)).content[0].text);
    const costs = { workbook: "Book", table: "Costs" };
    expect(
      await json("sort_table", {
        ...costs,
        column: "amount",
        direction: "descending",
      }),
    ).toMatchObject({ sort: "Amount descending", hiddenRows: [] });
    expect(engine.getCellValue(cellAddress("Book", "Sheet1", "A2"))).toBe("c");
    expect(
      await json("filter_table", {
        ...costs,
        column: "Amount",
        operator: "greaterThan",
        value: 1,
      }),
    ).toMatchObject({ filters: ["Amount > 1"], hiddenRows: [4] });
    expect(
      (await json("filter_table", { ...costs, column: "Item", values: ["A"] }))
        .hiddenRows,
    ).toEqual([2, 3, 4]);
    expect((await json("list_tables", { workbook: "Book" }))[0]).toMatchObject({
      name: "Costs",
      sheet: "Sheet1",
      range: "A1:B4",
      columns: ["Item", "Amount"],
      filters: ["Item: A", "Amount > 1"],
    });
    expect(
      (await json("filter_table", { ...costs, column: "Amount" })).hiddenRows,
    ).toEqual([2, 3]);
    expect(engine.getHiddenRows(sheet)).toEqual(new Set([1, 2]));
    for (const args of [
      { ...costs, column: "Missing" },
      { ...costs, column: "Item", values: ["a"], operator: "equals" },
      { ...costs, column: "Item", operator: "near", value: "a" },
    ])
      expect((await call("filter_table", args)).isError).toBe(true);
    expect(
      (await call("sort_table", { ...costs, column: "Item", direction: "up" }))
        .isError,
    ).toBe(true);
    expect((await call("list_tables", { workbook: "Nope" })).isError).toBe(
      true,
    );
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
