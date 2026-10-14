import { parseCellReference, type FormulaEngine } from "@ricsam/formula-engine";

export function cellAddress(
  workbookName: string,
  sheetName: string,
  cell: string,
) {
  if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/i.test(cell))
    throw new Error("Use a cell address such as A1 or D20.");
  const { rowIndex, colIndex } = parseCellReference(cell.toUpperCase());
  if (rowIndex >= 1048576 || colIndex >= 16384)
    throw new Error("Cell is outside Excel’s worksheet limits.");
  return { workbookName, sheetName, rowIndex, colIndex };
}

export function textValue(value: unknown): string {
  if (value === undefined || value === null) return "";
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

export function parseInput(
  value: string,
): string | number | boolean | undefined {
  if (value === "") return undefined;
  if (/^(true|false)$/i.test(value)) return value.toLowerCase() === "true";
  if (
    /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(value) &&
    Number.isFinite(Number(value))
  )
    return Number(value);
  return value;
}

export function inspectCell(
  engine: FormulaEngine,
  workbookName: string,
  sheetName: string,
  cell: string,
) {
  if (!engine.getSheet({ workbookName, sheetName }))
    throw new Error("Worksheet not found.");
  const address = cellAddress(workbookName, sheetName, cell);
  return {
    cell: cell.toUpperCase(),
    content:
      engine
        .getSheet({ workbookName, sheetName })
        ?.content.get(cell.toUpperCase()) ?? null,
    value: engine.getCellValue(address) ?? null,
    error: engine.getCellErrorDetails(address) ?? null,
  };
}
