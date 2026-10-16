import {
  FormulaEngine,
  getCellReference,
  parseCellReference,
  type CellAddress,
  type RangeAddress,
  type SerializedCellValue,
} from "@ricsam/formula-engine";
import { WorkbookClipboardManager } from "@ricsam/react-spreadsheets";
import { SelectionManager } from "@ricsam/selection-manager";

export type PasteFormatting = "with-formatting" | "without-formatting";
type CopyContext = Parameters<WorkbookClipboardManager["triggerCopy"]>[0];
type PasteContext = Parameters<WorkbookClipboardManager["triggerPaste"]>[0];
const MAX_CELLS = 100_000;

export function selectionForRanges(ranges: RangeAddress[]): SelectionManager {
  const sm = new SelectionManager(
    () => ({ type: "infinity" }),
    () => ({ type: "infinity" }),
    () => [],
  );
  sm.replaceSelections(
    ranges.map(({ range }) => ({
      start: range.start,
      end: {
        row:
          range.end.row.type === "number"
            ? range.end.row
            : { type: "infinity" },
        col:
          range.end.col.type === "number"
            ? range.end.col
            : { type: "infinity" },
      },
    })),
  );
  return sm;
}

/** Extends, rather than intercepts, FormulaSheet's keyboard clipboard callbacks. */
export class WorkspaceClipboardManager extends WorkbookClipboardManager {
  private snapshot?: string;
  private listeners = new Set<() => void>();
  private version = 0;
  constructor(
    private readonly workspaceEngine: FormulaEngine,
    private readonly isLocked: () => boolean = () => false,
  ) {
    super(workspaceEngine);
  }
  subscribe = (listener: () => void) => {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  };
  getVersion = () => this.version;
  get hasContent() {
    return !!this.snapshot && this.copiedCells.length > 0;
  }

  override triggerCopy(context: CopyContext): void {
    if (this.isLocked()) return;
    const sheet = this.workspaceEngine.getSheet(context);
    if (!sheet) return;
    const area = context.selectionManager.selections.at(-1);
    if (!area) return;
    const startRow = Math.min(
      area.start.row,
      area.end.row.type === "number" ? area.end.row.value : area.start.row,
    );
    const startCol = Math.min(
      area.start.col,
      area.end.col.type === "number" ? area.end.col.value : area.start.col,
    );
    let endRow =
      area.end.row.type === "number"
        ? Math.max(area.start.row, area.end.row.value)
        : startRow;
    let endCol =
      area.end.col.type === "number"
        ? Math.max(area.start.col, area.end.col.value)
        : startCol;
    const assertSize = () => {
      if ((endRow - startRow + 1) * (endCol - startCol + 1) > MAX_CELLS)
        throw new Error("Copy up to 100,000 cells at a time.");
    };
    assertSize();
    // Bound open-ended selections with a constant-space scan before extraction.
    if (area.end.row.type === "infinity" || area.end.col.type === "infinity") {
      for (const key of sheet.content.keys()) {
        const cell = parseCellReference(key);
        if (
          cell.rowIndex < startRow ||
          cell.colIndex < startCol ||
          (area.end.row.type === "number" && cell.rowIndex > endRow) ||
          (area.end.col.type === "number" && cell.colIndex > endCol)
        )
          continue;
        if (area.end.row.type === "infinity")
          endRow = Math.max(endRow, cell.rowIndex);
        if (area.end.col.type === "infinity")
          endCol = Math.max(endCol, cell.colIndex);
        assertSize();
      }
    }
    const extracted = this.extractCellsFromSelection(
      context.selectionManager,
      sheet.content,
    );
    if (!extracted) return;
    if (extracted.width * extracted.height > MAX_CELLS)
      throw new Error("Copy up to 100,000 cells at a time.");
    const selected = context.selectionManager.selections.at(-1)!;
    const row = Math.min(
      selected.start.row,
      selected.end.row.type === "number"
        ? selected.end.row.value
        : selected.start.row,
    );
    const col = Math.min(
      selected.start.col,
      selected.end.col.type === "number"
        ? selected.end.col.value
        : selected.start.col,
    );
    // Include blank cells: the upstream extraction only returns occupied cells,
    // which otherwise shifts ranges and loses blank-cell formatting.
    const cells: CellAddress[] = [];
    for (let r = 0; r < extracted.height; r++)
      for (let c = 0; c < extracted.width; c++)
        cells.push({
          workbookName: context.workbookName,
          sheetName: context.sheetName,
          rowIndex: row + r,
          colIndex: col + c,
        });
    const snapshot = this.workspaceEngine.serializeEngine({
      includeEvaluationState: false,
    });
    super.triggerCopy(context);
    this.copiedCells = cells;
    this.snapshot = snapshot;
    this.version++;
    this.listeners.forEach((listener) => listener());
  }

  copyRanges(ranges: RangeAddress[], cut = false): void {
    const source = ranges.at(-1);
    if (source)
      this.triggerCopy({
        ...source,
        selectionManager: selectionForRanges(ranges),
        copyType: "value",
        cut,
      });
  }

  override triggerPaste(context: PasteContext): void {
    if (this.isLocked()) return;
    if (this.hasContent && context.rawString === this.signature) {
      this.pasteRanges(
        context.selectionManager.selections.map((range) => ({
          workbookName: context.workbookName,
          sheetName: context.sheetName,
          range: {
            start: range.start,
            end: {
              row:
                range.end.row.type === "number"
                  ? range.end.row
                  : { type: "infinity", sign: "positive" },
              col:
                range.end.col.type === "number"
                  ? range.end.col
                  : { type: "infinity", sign: "positive" },
            },
          },
        })),
        "with-formatting",
        context.pasteType,
      );
    } else {
      // Keep native CSV/TSV parsing, data coercion and editing behavior untouched.
      context.selectionManager.saveCellValues(context.updates);
    }
  }

  pasteRanges(
    ranges: RangeAddress[],
    formatting: PasteFormatting = "with-formatting",
    type: "value" | "formula" = "formula",
  ): void {
    if (this.isLocked() || !this.snapshot || !ranges.length) return;
    const engine = this.workspaceEngine;
    const source = this.copiedCells;
    let minRow = Infinity,
      minCol = Infinity,
      maxRow = -Infinity,
      maxCol = -Infinity;
    for (const cell of source) {
      minRow = Math.min(minRow, cell.rowIndex);
      minCol = Math.min(minCol, cell.colIndex);
      maxRow = Math.max(maxRow, cell.rowIndex);
      maxCol = Math.max(maxCol, cell.colIndex);
    }
    const height = maxRow - minRow + 1;
    const width = maxCol - minCol + 1;
    const destinations = ranges.map((address) => {
      if (!engine.getSheet(address)) throw new Error("Worksheet not found.");
      const { range } = address;
      const targetHeight =
        range.end.row.type === "number"
          ? range.end.row.value - range.start.row + 1
          : 1;
      const targetWidth =
        range.end.col.type === "number"
          ? range.end.col.value - range.start.col + 1
          : 1;
      const fill =
        !this.isCut && (targetHeight > height || targetWidth > width);
      // Upstream fill is column-first: even a narrower target receives all seed
      // columns, but a shorter target receives only its requested rows.
      const h = fill ? targetHeight : height;
      const w = fill ? Math.max(width, targetWidth) : width;
      const { row, col } = range.start;
      if (
        ![row, col, targetHeight, targetWidth].every(Number.isInteger) ||
        row < 0 ||
        col < 0 ||
        targetHeight < 1 ||
        targetWidth < 1 ||
        row + h > 1_048_576 ||
        col + w > 16_384
      )
        throw new Error("Paste is outside Excel’s worksheet limits.");
      // Bound open-ended ranges explicitly; the library's fill helper otherwise
      // assigns them an unrelated 100-cell extent.
      const bounded = {
        ...address,
        range: {
          start: range.start,
          end: {
            row: { type: "number" as const, value: row + targetHeight - 1 },
            col: { type: "number" as const, value: col + targetWidth - 1 },
          },
        },
      };
      return { address: bounded, h, w };
    });
    if (destinations.reduce((n, d) => n + d.h * d.w, 0) > MAX_CELLS)
      throw new Error("Paste up to 100,000 cells at a time.");
    const include =
      formatting === "with-formatting"
        ? ("all" as const)
        : ["content" as const];
    if (this.isCut) {
      // A move must run against the live engine to update dependent references.
      if (ranges.length !== 1)
        throw new Error("Cut cells can only be pasted to one range.");
      if (source.some((cell) => !engine.getSheet(cell)))
        throw new Error("The cut source worksheet no longer exists.");
      engine.smartPaste(
        source,
        { ...ranges[0], areas: [ranges[0].range] },
        { cut: true, type, include },
      );
      this.isCut = false;
      this.snapshot = undefined;
      this.version++;
      this.listeners.forEach((listener) => listener());
      return;
    }
    // Use an isolated engine to reuse its formula translation/fill semantics.
    // Never temporarily mutate the user's source cells or restore the live engine.
    const frozen = new FormulaEngine();
    const snapshot = this.snapshot;
    engine.transact(() => {
      for (const { address, h, w } of destinations) {
        frozen.resetToSerializedEngine(snapshot);
        if (!frozen.getWorkbooks().has(address.workbookName))
          frozen.addWorkbook({
            workbookName: address.workbookName,
            data: { sheets: [{ name: address.sheetName }] },
          });
        else if (!frozen.getSheet(address)) frozen.addSheet(address);
        // Upstream skips undefined source content rather than clearing the
        // destination. Record blanks before overlap can mutate the source.
        const blanks = new Set<string>();
        for (const cell of source)
          if (
            frozen.getSheet(cell)?.content.get(getCellReference(cell)) ===
            undefined
          )
            blanks.add(`${cell.rowIndex - minRow}:${cell.colIndex - minCol}`);
        frozen.smartPaste(
          source,
          { ...address, areas: [address.range] },
          { type, include },
        );
        const content = frozen.getSheet(address)!.content;
        for (let r = 0; r < h; r++)
          for (let c = 0; c < w; c++) {
            const cell = {
              workbookName: address.workbookName,
              sheetName: address.sheetName,
              rowIndex: address.range.start.row + r,
              colIndex: address.range.start.col + c,
            };
            const value: SerializedCellValue = blanks.has(
              `${r % height}:${c % width}`,
            )
              ? undefined
              : content.get(getCellReference(cell));
            engine.setCellContent(cell, value);
            if (formatting === "with-formatting") {
              const range: RangeAddress = {
                ...address,
                range: {
                  start: { row: cell.rowIndex, col: cell.colIndex },
                  end: {
                    row: { type: "number", value: cell.rowIndex },
                    col: { type: "number", value: cell.colIndex },
                  },
                },
              };
              engine.setCellMetadata(
                cell,
                structuredClone(frozen.getCellMetadata(cell)),
              );
              engine.clearRangeMetadata(range);
              for (const metadata of frozen.getRangeMetadataForCell(cell))
                engine.addRangeMetadata({
                  areas: [range],
                  metadata: structuredClone(metadata.metadata),
                });
              engine.clearCellStyles(range);
              engine.clearCellDataTypes(range);
              const style = frozen.getCellStyle(cell);
              if (style)
                engine.addCellStyle({
                  areas: [range],
                  style: structuredClone(style),
                });
              engine.addCellDataType({
                areas: [range],
                dataType: frozen.getCellDataType(cell),
              });
            }
          }
      }
    });
  }
}
