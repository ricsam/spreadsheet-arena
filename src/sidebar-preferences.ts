import { useEffect, useState } from "react";

const STORAGE_KEY = "spreadsheet-arena.sidebar-widths";
export function readSidebarWidths(): { left: number; right: number } {
  try {
    const value = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "null");
    return {
      left:
        typeof value?.left === "number" && Number.isFinite(value.left)
          ? Math.max(220, Math.min(420, value.left))
          : 240,
      right:
        typeof value?.right === "number" && Number.isFinite(value.right)
          ? Math.max(260, Math.min(440, value.right))
          : 280,
    };
  } catch {
    return { left: 240, right: 280 };
  }
}

export function useSidebarPreferences() {
  const [widths, setWidths] = useState(readSidebarWidths);
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(widths));
    } catch {
      /* Optional device preference. */
    }
  }, [widths]);
  return {
    leftWidth: widths.left,
    rightWidth: widths.right,
    setLeftWidth: (left: number) => setWidths((value) => ({ ...value, left })),
    setRightWidth: (right: number) =>
      setWidths((value) => ({ ...value, right })),
  };
}
