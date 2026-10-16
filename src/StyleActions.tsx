import type { FormulaEngine } from "@ricsam/formula-engine";
import type { WorkbookSelectionManager } from "@ricsam/react-spreadsheets";
import { Bold, Italic, Underline } from "lucide-react";
import type { Pane } from "./model";
import { useActiveSelection } from "./workspace-selection";
import { applyStyle, firstAddress } from "./formatting";
import { Button } from "./components/ui/button";

export interface FormattingProps {
  engine: FormulaEngine;
  selection: WorkbookSelectionManager;
  active?: Pane;
  revision: number;
  isLocked?: () => boolean;
  onError: (error: unknown) => void;
}
export function StyleActions({
  engine,
  selection,
  active,
  isLocked = () => false,
  onError,
}: FormattingProps) {
  const areas = useActiveSelection(selection, active);
  const address = firstAddress(areas);
  const style = address ? engine.getCellStyle(address) : undefined;
  const disabled = !active || !areas.length;
  const run = (action: () => void) => {
    try {
      if (!isLocked()) action();
    } catch (error) {
      onError(error);
    }
  };
  return (
    <div
      className="style-actions flex items-center gap-1"
      role="group"
      aria-label="Quick cell formatting"
    >
      {(
        [
          ["bold", "Bold", Bold],
          ["italic", "Italic", Italic],
          ["underline", "Underline", Underline],
        ] as const
      ).map(([key, label, Icon]) => (
        <Button
          key={key}
          type="button"
          size="icon"
          variant="ghost"
          className="size-7 aria-pressed:bg-accent aria-pressed:text-accent-foreground"
          aria-label={label}
          title={label}
          aria-pressed={Boolean(style?.[key])}
          disabled={disabled}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() =>
            run(() => applyStyle(engine, areas, { [key]: !style?.[key] }))
          }
        >
          <Icon size={15} />
        </Button>
      ))}
      <input
        className="h-7 w-7 cursor-pointer rounded border p-0.5"
        type="color"
        aria-label="Text color"
        title="Text color"
        disabled={disabled}
        value={
          /^#[\da-f]{6}$/i.test(style?.color ?? "") ? style!.color : "#172b36"
        }
        onChange={(e) =>
          run(() => applyStyle(engine, areas, { color: e.target.value }))
        }
      />
      <input
        className="h-7 w-7 cursor-pointer rounded border p-0.5"
        type="color"
        aria-label="Fill color"
        title="Fill color"
        disabled={disabled}
        value={
          /^#[\da-f]{6}$/i.test(style?.backgroundColor ?? "")
            ? style!.backgroundColor
            : "#ffffff"
        }
        onChange={(e) =>
          run(() =>
            applyStyle(engine, areas, { backgroundColor: e.target.value }),
          )
        }
      />
    </div>
  );
}
