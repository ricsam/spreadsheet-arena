import { useMemo, type ComponentProps } from "react";
import { FormulaSheet } from "@ricsam/react-spreadsheets";

export function WorkbookSheet({
  theme,
  ...props
}: ComponentProps<typeof FormulaSheet> & { theme: "light" | "dark" }) {
  // The library's canvas redraws on dimension changes, but not inherited CSS
  // theme changes. Refresh the dimension identity without remounting the grid:
  // selection, scroll position, and in-progress cell editors stay intact.
  const columnWidths = useMemo(
    () => ({ ...props.columnWidths }),
    [props.columnWidths, theme],
  );
  return <FormulaSheet {...props} columnWidths={columnWidths} />;
}
