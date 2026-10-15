import { useCallback, useMemo, useSyncExternalStore } from "react";
import { type RangeAddress } from "@ricsam/formula-engine";
import { type WorkbookSelectionManager } from "@ricsam/react-spreadsheets";
import { cellAddress } from "./cells";
import type { Pane } from "./model";

/** Toolbar focus is not grid focus. Retain the last range only for this exact view. */
export function getActiveSelection(
  selection: WorkbookSelectionManager,
  active?: Pane,
): RangeAddress[] {
  if (!active) return [];
  const context = selection.getLastFocusedSelectionManager()?.context;
  if (
    context?.workbookName === active.workbookName &&
    context.sheetName === active.sheetName &&
    context.viewId === active.id
  ) {
    const ranges = selection.getSelections();
    if (ranges.length) return ranges;
  }
  const { rowIndex: row, colIndex: col } = cellAddress(
    active.workbookName,
    active.sheetName,
    active.cell ?? "A1",
  );
  return [
    {
      workbookName: active.workbookName,
      sheetName: active.sheetName,
      range: {
        start: { row, col },
        end: {
          row: { type: "number", value: row },
          col: { type: "number", value: col },
        },
      },
    },
  ];
}

export function useActiveSelection(
  selection: WorkbookSelectionManager,
  active?: Pane,
): RangeAddress[] {
  const subscribe = useCallback(
    (notify: () => void) => selection.onSelectionChange(notify),
    [selection],
  );
  // A primitive snapshot is stable between renders, including range-only changes.
  const snapshot = useCallback(
    () => JSON.stringify(getActiveSelection(selection, active)),
    [selection, active],
  );
  const json = useSyncExternalStore(subscribe, snapshot, snapshot);
  return useMemo(() => JSON.parse(json) as RangeAddress[], [json]);
}
