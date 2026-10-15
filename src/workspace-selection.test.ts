import { describe, expect, it } from "vitest";
import { WorkbookSelectionManager } from "@ricsam/react-spreadsheets";
import { selectionForRanges } from "./workspace-clipboard";
import { getActiveSelection } from "./workspace-selection";
import type { RangeAddress } from "@ricsam/formula-engine";
const active = {
  id: "view1",
  workbookName: "Book",
  sheetName: "Sheet",
  cell: "C3",
};
const range: RangeAddress = {
  workbookName: "Book",
  sheetName: "Sheet",
  range: {
    start: { row: 0, col: 0 },
    end: {
      row: { type: "number", value: 4 },
      col: { type: "number", value: 2 },
    },
  },
};
describe("active selection", () => {
  it("retains ranges on grid blur but never targets another workbook or duplicate view", () => {
    const selection = new WorkbookSelectionManager();
    const sm = selectionForRanges([range]);
    selection.add(sm, { ...active, viewId: active.id });
    sm.focus();
    expect(getActiveSelection(selection, active)).toEqual([range]);
    sm.blur();
    expect(getActiveSelection(selection, active)).toEqual([range]);
    expect(
      getActiveSelection(selection, { ...active, id: "view2" })[0].range.start,
    ).toEqual({ row: 2, col: 2 });
    expect(
      getActiveSelection(selection, { ...active, workbookName: "Other" })[0]
        .workbookName,
    ).toBe("Other");
    expect(getActiveSelection(selection)).toEqual([]);
  });
  it("emits range-only changes and falls back after manager unregister", () => {
    const selection = new WorkbookSelectionManager();
    const sm = selectionForRanges([range]);
    const remove = selection.add(sm, { ...active, viewId: active.id });
    sm.focus();
    let changes = 0;
    const unsubscribe = selection.onSelectionChange(() => changes++);
    sm.replaceSelections([
      {
        ...sm.selections[0],
        end: {
          row: { type: "number", value: 9 },
          col: { type: "number", value: 9 },
        },
      },
    ]);
    expect(changes).toBe(1);
    expect(getActiveSelection(selection, active)[0].range.end.row).toEqual({
      type: "number",
      value: 9,
    });
    unsubscribe();
    remove();
    expect(getActiveSelection(selection, active)[0].range.start).toEqual({
      row: 2,
      col: 2,
    });
  });
});
