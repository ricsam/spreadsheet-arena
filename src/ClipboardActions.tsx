import { useSyncExternalStore } from "react";
import type { FormulaEngine } from "@ricsam/formula-engine";
import type { WorkbookSelectionManager } from "@ricsam/react-spreadsheets";
import { Copy, Scissors, ClipboardPaste, ChevronDown } from "lucide-react";
import { Button } from "./components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "./components/ui/dropdown-menu";
import type { Pane } from "./model";
import { parseInput } from "./cells";
import { useActiveSelection } from "./workspace-selection";
import {
  WorkspaceClipboardManager,
  selectionForRanges,
  type PasteFormatting,
} from "./workspace-clipboard";

export function ClipboardActions({
  engine,
  selection,
  clipboard,
  active,
  isLocked = () => false,
  onError,
}: {
  engine: FormulaEngine;
  selection: WorkbookSelectionManager;
  clipboard: WorkspaceClipboardManager;
  active?: Pane;
  isLocked?: () => boolean;
  onError: (error: unknown) => void;
}) {
  const ranges = useActiveSelection(selection, active);
  useSyncExternalStore(
    clipboard.subscribe,
    clipboard.getVersion,
    clipboard.getVersion,
  );
  const run = (action: () => void) => {
    try {
      if (!isLocked()) action();
    } catch (error) {
      onError(error);
    }
  };
  const paste = async (formatting: PasteFormatting) => {
    // Capture the target before the asynchronous OS read; do not refocus a stale view.
    const targets = ranges;
    try {
      if (isLocked() || !targets.length) return;
      let text: string;
      try {
        text = await navigator.clipboard.readText();
      } catch (error) {
        // Internal copy remains usable when the browser denies clipboard access.
        if (isLocked()) return;
        if (!clipboard.hasContent) throw error;
        clipboard.pasteRanges(targets, formatting);
        return;
      }
      if (isLocked()) return;
      if (clipboard.hasContent && text === clipboard.signature) {
        clipboard.pasteRanges(targets, formatting);
        return;
      }
      const sm = selectionForRanges(targets);
      sm.focus();
      sm.listenToPaste(({ updates }) =>
        engine.transact(() => {
          for (const update of updates)
            engine.setCellContent(
              {
                workbookName: targets[0].workbookName,
                sheetName: targets[0].sheetName,
                rowIndex: update.rowIndex,
                colIndex: update.colIndex,
              },
              parseInput(update.value),
            );
        }),
      );
      // Reuse the library's quoted CSV/TSV parser, without a document listener.
      sm.handlePaste({
        preventDefault() {},
        clipboardData: { getData: () => text },
      } as unknown as ClipboardEvent);
    } catch (error) {
      onError(error);
    }
  };
  return (
    <div
      className="flex items-center gap-1"
      role="group"
      aria-label="Clipboard"
    >
      <Button
        variant="ghost"
        size="icon"
        aria-label="Cut cells"
        title="Cut cells"
        disabled={!ranges.length}
        onClick={() => run(() => clipboard.copyRanges(ranges, true))}
      >
        <Scissors className="size-4" />
      </Button>
      <Button
        variant="ghost"
        size="icon"
        aria-label="Copy cells"
        title="Copy cells"
        disabled={!ranges.length}
        onClick={() => run(() => clipboard.copyRanges(ranges))}
      >
        <Copy className="size-4" />
      </Button>
      <Button
        variant="ghost"
        size="sm"
        disabled={!ranges.length}
        onClick={() => void paste("with-formatting")}
      >
        <ClipboardPaste className="size-4" />
        Paste
      </Button>
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <Button
            variant="ghost"
            size="icon"
            aria-label="Paste options"
            disabled={!ranges.length}
          >
            <ChevronDown className="size-3" />
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start">
          <DropdownMenuItem onSelect={() => void paste("with-formatting")}>
            Paste with formatting
          </DropdownMenuItem>
          <DropdownMenuItem onSelect={() => void paste("without-formatting")}>
            Paste without formatting
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
    </div>
  );
}
