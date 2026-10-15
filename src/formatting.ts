import {
  indexToColumn,
  type FormulaEngine,
  type CellStyle,
  type CellDataType,
  type RangeAddress,
  type TableDefinition,
} from "@ricsam/formula-engine";

export const SHEET_ROWS = 1048576;
export const SHEET_COLUMNS = 16384;

/** Styles remain sparse range rules, including whole rows and columns. */
export function validateRanges(engine: FormulaEngine, areas: RangeAddress[]) {
  if (!areas.length || areas.length > 256)
    throw new Error("Invalid selection.");
  for (const area of areas) {
    if (!engine.getSheet(area)) throw new Error("Worksheet not found.");
    for (const [axis, limit] of [
      ["row", SHEET_ROWS],
      ["col", SHEET_COLUMNS],
    ] as const) {
      const start = area.range.start[axis],
        end = area.range.end[axis];
      if (
        !Number.isInteger(start) ||
        start < 0 ||
        start >= limit ||
        (end.type === "number" &&
          (!Number.isInteger(end.value) ||
            end.value < start ||
            end.value >= limit))
      )
        throw new Error("Selection is outside worksheet limits.");
    }
  }
}
export function applyStyle(
  engine: FormulaEngine,
  areas: RangeAddress[],
  style: CellStyle,
) {
  validateRanges(engine, areas);
  if (
    style.fontSize !== undefined &&
    (!Number.isFinite(style.fontSize) ||
      style.fontSize < 6 ||
      style.fontSize > 144)
  )
    throw new Error("Font size must be between 6 and 144.");
  engine.addCellStyle({ areas, style });
}
export function applyDataType(
  engine: FormulaEngine,
  areas: RangeAddress[],
  dataType: CellDataType,
) {
  validateRanges(engine, areas);
  if (!["general", "text", "number", "boolean"].includes(dataType))
    throw new Error("Invalid data type.");
  engine.addCellDataType({ areas, dataType });
}
export function firstAddress(areas: RangeAddress[]) {
  const area = areas[0];
  return (
    area && {
      workbookName: area.workbookName,
      sheetName: area.sheetName,
      rowIndex: area.range.start.row,
      colIndex: area.range.start.col,
    }
  );
}
export function finiteTableRange(engine: FormulaEngine, areas: RangeAddress[]) {
  validateRanges(engine, areas);
  if (areas.length !== 1)
    throw new Error("A table requires one rectangular selection.");
  const area = areas[0],
    { start, end } = area.range;
  if (end.row.type !== "number" || end.col.type !== "number")
    throw new Error("A table requires a finite range.");
  const rows = end.row.value - start.row + 1,
    columns = end.col.value - start.col + 1;
  if (rows * columns > 100000 || columns > 1000)
    throw new Error("Table range is too large.");
  return { area, rows, columns, endRow: end.row.value, endCol: end.col.value };
}
export function validateTableName(
  engine: FormulaEngine,
  workbookName: string,
  name: string,
  previous?: string,
) {
  if (
    !/^[A-Za-z_][A-Za-z0-9_]{0,254}$/.test(name) ||
    /^[A-Za-z]{1,3}[1-9]\d*$/.test(name) ||
    /^[RC]$/i.test(name)
  )
    throw new Error("Invalid table name.");
  if (
    [...engine.getTables(workbookName).values()].some(
      (t) => t.name !== previous && t.name.toLowerCase() === name.toLowerCase(),
    )
  )
    throw new Error("Table name already exists.");
}
export function saveTable(
  engine: FormulaEngine,
  areas: RangeAddress[],
  name: string,
  previous?: TableDefinition,
) {
  const { area, rows, columns, endRow, endCol } = finiteTableRange(
    engine,
    areas,
  );
  name = name.trim();
  validateTableName(engine, area.workbookName, name, previous?.name);
  for (const table of engine.getTables(area.workbookName).values()) {
    if (table.name === previous?.name || table.sheetName !== area.sheetName)
      continue;
    const bottom =
      table.endRow.type === "number" ? table.endRow.value : Infinity;
    if (
      area.range.start.row <= bottom &&
      endRow >= table.start.rowIndex &&
      area.range.start.col <= table.start.colIndex + table.headers.size - 1 &&
      endCol >= table.start.colIndex
    )
      throw new Error("Tables cannot overlap.");
  }
  const options = {
    workbookName: area.workbookName,
    sheetName: area.sheetName,
    start: `${indexToColumn(area.range.start.col)}${area.range.start.row + 1}`,
    numRows: { type: "number" as const, value: rows - 1 },
    numCols: columns,
  };
  engine.transact(() => {
    if (previous) {
      engine.updateTable({ ...options, tableName: previous.name });
      if (name !== previous.name)
        engine.renameTable(area.workbookName, {
          oldName: previous.name,
          newName: name,
        });
    } else engine.addTable({ ...options, tableName: name });
  });
}
export function tableRange(table: TableDefinition): RangeAddress {
  return {
    workbookName: table.workbookName,
    sheetName: table.sheetName,
    range: {
      start: { row: table.start.rowIndex, col: table.start.colIndex },
      end: {
        row: table.endRow,
        col: {
          type: "number",
          value: table.start.colIndex + table.headers.size - 1,
        },
      },
    },
  };
}
export function setDimension(
  engine: FormulaEngine,
  areas: RangeAddress[],
  axis: "row" | "col",
  size: number,
) {
  validateRanges(engine, areas);
  if (
    !Number.isFinite(size) ||
    size < 16 ||
    size > (axis === "row" ? 500 : 1000)
  )
    throw new Error("Invalid dimension.");
  const indexes = new Set<number>();
  const first = areas[0];
  for (const area of areas) {
    if (
      area.workbookName !== first.workbookName ||
      area.sheetName !== first.sheetName
    )
      throw new Error("Select one worksheet.");
    const end = area.range.end[axis];
    if (end.type !== "number" || end.value - area.range.start[axis] > 10000)
      throw new Error("Dimension selection is too large.");
    for (let i = area.range.start[axis]; i <= end.value; i++) indexes.add(i);
    if (indexes.size > 10000)
      throw new Error("Dimension selection is too large.");
  }
  const old = engine.getSheetMetadata(first) as
    Record<string, unknown> | undefined;
  const key = axis === "row" ? "arenaRowHeights" : "arenaColumnWidths";
  const dimensions = { ...(old?.[key] as Record<string, number> | undefined) };
  for (const i of indexes)
    dimensions[axis === "row" ? String(i + 1) : indexToColumn(i)] = size;
  engine.setSheetMetadata(first, { ...old, [key]: dimensions });
}
