import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { FormulaEngine, indexToColumn } from "@ricsam/formula-engine";
import { WorkbookSelectionManager } from "@ricsam/react-spreadsheets";
import { WorkbookSheet } from "./WorkbookSheet";
import {
  ArrowDownToLine,
  ArrowUpFromLine,
  Bug,
  Check,
  ChevronDown,
  Columns2,
  FileSpreadsheet,
  FolderOpen,
  HelpCircle,
  Maximize2,
  Minimize2,
  PanelLeftClose,
  PanelRightClose,
  PanelRightOpen,
  Moon,
  Sun,
  Plus,
  Redo2,
  RotateCcw,
  Rows2,
  Search,
  Save,
  Undo2,
  X,
} from "lucide-react";
import {
  closePane,
  createDemoEngine,
  createEmptyWorkbook,
  exportWorkbook,
  importFiles,
  makeSnapshot,
  panes,
  reconcileLayout,
  resizeSplit,
  splitPane,
  updatePane,
  restoreSnapshot,
  MAX_WORKSPACE_FILE_BYTES,
  type Layout,
  type Pane,
} from "./model";
import { loadWorkspace, createWorkspaceWriter } from "./persistence";
import { cellAddress, parseInput, textValue } from "./cells";
import { Tiles } from "./Tiles";
import { registerArenaTools } from "./webmcp";
import { Button } from "./components/ui/button";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuCheckboxItem,
} from "./components/ui/dropdown-menu";
import { ResizableSidebar } from "./ResizableSidebar";
import { useTheme } from "./theme";
import { useSidebarPreferences } from "./sidebar-preferences";
import { WorkspaceClipboardManager } from "./workspace-clipboard";
import { ClipboardActions } from "./ClipboardActions";
import { WorkbookInspector } from "./WorkbookInspector";
import { StyleActions } from "./StyleActions";

type Notice = { message: string; kind: "info" | "error" };
type DialogState = {
  title: string;
  label?: string;
  initial?: string;
  description: string;
  action: (value: string) => void;
  danger?: boolean;
};
function download(name: string, content: string) {
  const url = URL.createObjectURL(
    new Blob([content], { type: "application/json" }),
  );
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function IconButton({
  label,
  children,
  ...props
}: {
  label: string;
  children: ReactNode;
} & React.ButtonHTMLAttributes<HTMLButtonElement>) {
  return (
    <Button
      type="button"
      variant="ghost"
      size="icon"
      className="icon-button size-7 p-1"
      aria-label={label}
      title={label}
      {...props}
    >
      {children}
    </Button>
  );
}
function PromptDialog({
  state,
  onClose,
}: {
  state: DialogState;
  onClose: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [value, setValue] = useState(state.initial ?? "");
  const [error, setError] = useState("");
  useEffect(() => {
    dialog.current?.showModal();
  }, []);
  return (
    <dialog ref={dialog} onCancel={onClose} aria-labelledby="dialog-title">
      <form
        onSubmit={(event) => {
          event.preventDefault();
          try {
            state.action(value);
            onClose();
          } catch (e) {
            setError(String(e instanceof Error ? e.message : e));
          }
        }}
      >
        <div className="dialog-heading">
          <h2 id="dialog-title">{state.title}</h2>
          <IconButton label="Close dialog" onClick={onClose}>
            <X size={18} />
          </IconButton>
        </div>
        {state.description && <p>{state.description}</p>}
        {state.label && (
          <label className="field-label">
            {state.label}
            <input
              autoFocus
              required
              maxLength={100}
              value={value}
              onChange={(e) => setValue(e.target.value)}
            />
          </label>
        )}
        {error && (
          <p role="alert" className="error-text">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button type="button" className="button" onClick={onClose}>
            Cancel
          </button>
          <button className={`button ${state.danger ? "danger" : "primary"}`}>
            {state.danger ? "Remove workbook" : "Confirm"}
          </button>
        </div>
      </form>
    </dialog>
  );
}

function hasInspectorDraft() {
  const input = document.activeElement;
  return (
    input instanceof HTMLInputElement &&
    Boolean(input.closest(".workbook-inspector")) &&
    input.value !== input.defaultValue
  );
}

export function App() {
  const [engine, setEngine] = useState(() => FormulaEngine.buildEmpty());
  const [storedLayout, setLayout] = useState<Layout | null>(null);
  const [activePaneId, setActivePaneId] = useState<string | null>(null);
  const [revision, setRevision] = useState(0);
  const [ready, setReady] = useState(false);
  const [saveInProgress, setSaveInProgress] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [storageAvailable, setStorageAvailable] = useState(true);
  const [includeEvaluationState, setIncludeEvaluationState] = useState(false);
  const [savedState, setSavedState] = useState<{
    engine: FormulaEngine;
    revision: number;
    viewKey: string;
  } | null>(null);
  const [notice, setNotice] = useState<Notice | null>(null);
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [sidebar, setSidebar] = useState(
    () => !window.matchMedia("(max-width: 700px)").matches,
  );
  const [query, setQuery] = useState("");
  const [dialog, setDialog] = useState<DialogState | null>(null);
  const [help, setHelp] = useState(false);
  const [maximized, setMaximized] = useState<string | null>(null);
  const [showFormulas, setShowFormulas] = useState(false);
  // Debug tools are opt-in for this page session, not part of workbook backups.
  const [debugMode, setDebugMode] = useState(false);
  const { theme, toggleTheme } = useTheme();
  const { leftWidth, rightWidth, setLeftWidth, setRightWidth } =
    useSidebarPreferences();
  const [inspector, setInspector] = useState(
    () => window.matchMedia("(min-width: 1100px)").matches,
  );
  const [cell, setCell] = useState("A1");
  const [formula, setFormula] = useState("");
  const [formulaDirty, setFormulaDirty] = useState(false);
  const upload = useRef<HTMLInputElement>(null);
  const restore = useRef<HTMLInputElement>(null);
  const formulaInput = useRef<HTMLInputElement>(null);
  const dragDepth = useRef(0);
  const importInFlight = useRef(false);
  const saveLocked = useRef(false);
  const loadedBaseline = useRef(false);
  const saveDialog = useRef<HTMLDialogElement>(null);
  const [workspaceWriter, setWorkspaceWriter] = useState(createWorkspaceWriter);
  const selection = useMemo(() => new WorkbookSelectionManager(), [engine]);
  const clipboard = useMemo(
    () => new WorkspaceClipboardManager(engine, () => saveLocked.current),
    [engine],
  );
  // Engine updates are synchronous, but effects run after render. Reconcile
  // here so deletion/undo never renders a FormulaSheet for a missing workbook.
  // Commit the pruned tree too: undo restores data, not previously closed panes.
  const layout = reconcileLayout(storedLayout, engine);
  if (layout !== storedLayout) setLayout(layout);
  const allPanes = panes(layout);
  const active = allPanes.find((p) => p.id === activePaneId) ?? allPanes[0];
  const books = [...engine.getWorkbooks().keys()];
  const notifyError = useCallback(
    (e: unknown) =>
      setNotice({
        kind: "error",
        message: e instanceof Error ? e.message : String(e),
      }),
    [],
  );
  const patchPane = useCallback(
    (id: string, patch: Partial<Pane>) =>
      setLayout((current) => current && updatePane(current, id, patch)),
    [],
  );

  useEffect(() => {
    let cancelled = false;
    loadWorkspace()
      .then((loaded) => {
        const saved = loaded?.snapshot;
        if (cancelled) return;
        if (saved) {
          setWorkspaceWriter(createWorkspaceWriter(loaded!.storageToken));
          loadedBaseline.current = true;
          setEngine(loaded!.engine);
          setLayout(saved.layout);
          setActivePaneId(saved.activePaneId);
          if (saved.view) {
            setSidebar(saved.view.sidebar);
            setShowFormulas(saved.view.showFormulas);
            setMaximized(saved.view.maximized);
          }
        } else {
          const demo = createDemoEngine();
          setEngine(demo);
          const first: Pane = {
            id: crypto.randomUUID(),
            workbookName: "Launch budget",
            sheetName: "Overview",
            cell: "D12",
          };
          const second: Pane = {
            id: crypto.randomUUID(),
            workbookName: "Assumptions",
            sheetName: "Inputs",
            cell: "B4",
          };
          setLayout({
            kind: "split",
            id: crypto.randomUUID(),
            direction: "horizontal",
            ratio: 0.64,
            first: { kind: "pane", pane: first },
            second: { kind: "pane", pane: second },
          });
          setActivePaneId(first.id);
        }
        setReady(true);
      })
      .catch((e) => {
        if (!cancelled) {
          setStorageAvailable(false);
          setReady(true);
          notifyError(
            `Could not restore local data. It has not been overwritten. ${e.message}`,
          );
        }
      });
    return () => {
      cancelled = true;
    };
  }, [notifyError]);

  useEffect(() => {
    const narrow = window.matchMedia("(max-width: 900px)");
    const collapsePanels = () => {
      if (narrow.matches) {
        setSidebar(false);
        setInspector(false);
      }
    };
    narrow.addEventListener("change", collapsePanels);
    return () => narrow.removeEventListener("change", collapsePanels);
  }, []);

  useEffect(() => engine.onUpdate(() => setRevision((r) => r + 1)), [engine]);
  const workspaceView = {
    layout,
    activePaneId: active?.id ?? null,
    view: {
      sidebar,
      showFormulas,
      maximized: allPanes.some((p) => p.id === maximized) ? maximized : null,
    },
  };
  // Small view metadata only: never serialize engine data to detect dirty state.
  const viewKey = JSON.stringify(workspaceView);
  useEffect(() => {
    if (ready && loadedBaseline.current) {
      loadedBaseline.current = false;
      setSavedState({ engine, revision, viewKey });
    }
  }, [ready, engine, revision, viewKey]);
  const unsaved =
    !savedState ||
    savedState.engine !== engine ||
    savedState.revision !== revision ||
    savedState.viewKey !== viewKey;
  const saving = saveInProgress
    ? "Saving…"
    : !storageAvailable
      ? "Saving unavailable"
      : saveError
        ? "Save failed — unsaved changes"
        : unsaved || formulaDirty
          ? "Unsaved changes"
          : "Saved on this device";
  const isLocked = useCallback(() => saveLocked.current, []);
  const saveWorkspace = async () => {
    if (
      saveLocked.current ||
      importInFlight.current ||
      !ready ||
      !storageAvailable ||
      dialog
    )
      return;
    const editing =
      selection.getLastFocusedSelectionManager()?.sm.isEditing.type ?? "none";
    if (formulaDirty || editing !== "none" || hasInspectorDraft()) {
      notifyError(
        "Apply or cancel the cell edit before saving (Enter or Escape).",
      );
      return;
    }
    saveLocked.current = true;
    setSaveInProgress(true);
    setSaveError(false);
    try {
      // Paint the lock/notice first. Serialization is deliberately synchronous;
      // the lock remains in force until the atomic IndexedDB write completes.
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => setTimeout(resolve, 0)),
      );
      await workspaceWriter.save(engine, workspaceView, {
        includeEvaluationState,
      });
      setSavedState({ engine, revision, viewKey });
    } catch (error) {
      setSaveError(true);
      notifyError(
        `Local save failed. Your changes are still open; retry or download a backup. ${error instanceof Error ? error.message : error}`,
      );
    } finally {
      saveLocked.current = false;
      setSaveInProgress(false);
    }
  };
  useEffect(() => {
    if (saveInProgress) saveDialog.current?.showModal();
  }, [saveInProgress]);
  useEffect(() => {
    const keys = (event: KeyboardEvent) => {
      if (saveLocked.current) {
        event.preventDefault();
        event.stopImmediatePropagation();
      } else if (
        (event.ctrlKey || event.metaKey) &&
        event.key.toLowerCase() === "s" &&
        !event.altKey
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        void saveWorkspace();
      }
    };
    const block = (event: Event) => {
      if (saveLocked.current) {
        event.preventDefault();
        event.stopImmediatePropagation();
      }
    };
    window.addEventListener("keydown", keys, true);
    const events = [
      "pointerdown",
      "pointerup",
      "pointermove",
      "mousedown",
      "mouseup",
      "mousemove",
      "click",
      "keyup",
      "paste",
      "cut",
      "drop",
      "beforeinput",
    ];
    events.forEach((name) => window.addEventListener(name, block, true));
    return () => {
      window.removeEventListener("keydown", keys, true);
      events.forEach((name) => window.removeEventListener(name, block, true));
    };
  });

  useEffect(
    () =>
      selection.onSelectionChange(() => {
        const focused = selection.getLastFocusedSelectionManager();
        const selected = selection.getLastSelection();
        // Blur/mouseup can notify cached selections from every mounted view.
        // Only the keyboard-active grid should replace a formula-bar draft.
        if (!focused || !selected || !focused.sm.hasFocus) return;
        const id = focused.context.viewId;
        const address = `${indexToColumn(selected.range.start.col)}${selected.range.start.row + 1}`;
        if (id) {
          setActivePaneId(id);
          patchPane(id, { cell: address });
        }
      }),
    [selection, patchPane],
  );

  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (
        saveLocked.current ||
        unsaved ||
        saveError ||
        formulaDirty ||
        hasInspectorDraft() ||
        busy ||
        (selection.getLastFocusedSelectionManager()?.sm.isEditing.type ??
          "none") !== "none"
      ) {
        event.preventDefault();
        event.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [unsaved, saveError, formulaDirty, busy, selection]);

  const raw = active
    ? textValue(engine.getSheet(active)?.content.get(active.cell ?? "A1"))
    : "";
  useEffect(() => {
    setCell(active?.cell ?? "A1");
    setFormula(raw);
    setFormulaDirty(false);
  }, [active?.id, active?.sheetName, active?.cell, raw]);

  const focusCell = useCallback(
    (pane: Pane, address: string) => {
      const parsed = cellAddress(pane.workbookName, pane.sheetName, address);
      setActivePaneId(pane.id);
      patchPane(pane.id, { cell: address.toUpperCase() });
      selection.focusCell(
        pane.workbookName,
        pane.sheetName,
        parsed.rowIndex,
        parsed.colIndex,
        { viewId: pane.id, align: "nearest" },
      );
    },
    [selection, patchPane],
  );

  const openBook = useCallback(
    (workbookName: string, sheetName?: string) => {
      const existing = panes(layout).find(
        (p) => p.workbookName === workbookName,
      );
      const targetSheet =
        sheetName ??
        existing?.sheetName ??
        engine.getOrderedSheetNames(workbookName)[0];
      if (!targetSheet) throw new Error("Workbook has no sheets.");
      if (existing) {
        patchPane(existing.id, { sheetName: targetSheet });
        setActivePaneId(existing.id);
        setMaximized(null);
        return { ...existing, sheetName: targetSheet };
      }
      const pane: Pane = {
        id: crypto.randomUUID(),
        workbookName,
        sheetName: targetSheet,
        cell: "A1",
      };
      if (!layout) setLayout({ kind: "pane", pane });
      else if (panes(layout).length < 8)
        setLayout(
          splitPane(
            layout,
            active?.id ?? panes(layout)[0].id,
            pane,
            "horizontal",
          ),
        );
      else
        throw new Error(
          "Eight panes are open. Close a pane before opening another.",
        );
      setActivePaneId(pane.id);
      setMaximized(null);
      return pane;
    },
    [layout, active?.id, engine, patchPane],
  );

  const createBook = useCallback(
    (name: string) => {
      if (saveLocked.current)
        throw new Error("Saving in progress. Editing is temporarily disabled.");
      if (!name.trim()) throw new Error("Give your workbook a name.");
      if (panes(layout).length >= 8)
        throw new Error("Close a pane before creating another workbook.");
      const actual = createEmptyWorkbook(engine, name.trim());
      openBook(actual);
      return actual;
    },
    [engine, openBook, layout],
  );

  // File parsing is asynchronous; open its result against the latest pane layout.
  const importTarget = useRef({ openBook, paneCount: allPanes.length });
  importTarget.current = { openBook, paneCount: allPanes.length };
  const importDropped = async (files: File[]) => {
    if (saveLocked.current || importInFlight.current || !files.length) return;
    importInFlight.current = true;
    setBusy(true);
    try {
      const result = await importFiles(engine, files);
      if (result.names.length && importTarget.current.paneCount < 8)
        importTarget.current.openBook(result.names[0]);
      else if (result.names.length)
        result.warnings.push(
          "All eight panes are in use. Close a pane, then open the imported workbook from the sidebar.",
        );
      setNotice({
        kind: result.errors.length ? "error" : "info",
        message:
          [
            result.names.length ? `Imported ${result.names.join(", ")}.` : "",
            ...result.errors,
            ...result.warnings,
          ]
            .filter(Boolean)
            .join(" ") || "No workbooks were imported.",
      });
    } catch (e) {
      notifyError(e);
    } finally {
      importInFlight.current = false;
      setBusy(false);
    }
  };

  const split = (pane: Pane, direction: "horizontal" | "vertical") => {
    if (allPanes.length >= 8)
      return notifyError(
        "Eight panes are open. Close a pane before splitting again.",
      );
    const copy = { ...pane, id: crypto.randomUUID() };
    setLayout(
      (current) => current && splitPane(current, pane.id, copy, direction),
    );
    setActivePaneId(copy.id);
    setMaximized(null);
  };
  const close = (id: string) => {
    setLayout((current) => current && closePane(current, id));
    if (maximized === id) setMaximized(null);
  };
  const downloadWorkspace = () => {
    try {
      download(
        "spreadsheet-arena.workspace.json",
        JSON.stringify({
          ...makeSnapshot(engine, layout, active?.id ?? null),
          view: {
            sidebar,
            showFormulas,
            maximized: panes(layout).some((p) => p.id === maximized)
              ? maximized
              : null,
          },
        }),
      );
    } catch (e) {
      notifyError(e);
    }
  };
  const clearFormulaCache = () => {
    if (saveLocked.current || !ready || busy || !books.length) return;
    try {
      // Keep the engine, workbooks, history, and mounted spreadsheet views intact.
      // The API notifies subscribers; visible cells then recalculate on demand.
      engine.clearCache();
      setNotice({
        kind: "info",
        message:
          "Formula cache cleared. Results recalculate on demand; workbook data and undo history are unchanged.",
      });
    } catch (e) {
      notifyError(e);
    }
  };
  const commitFormula = () => {
    if (saveLocked.current || !active) return;
    try {
      engine.setCellContent(
        cellAddress(active.workbookName, active.sheetName, active.cell ?? "A1"),
        parseInput(formula),
      );
      setFormulaDirty(false);
      focusCell(active, active.cell ?? "A1");
    } catch (e) {
      notifyError(e);
    }
  };

  useEffect(
    () =>
      registerArenaTools({
        engine,
        layout,
        isLocked,
        activePaneId: active?.id ?? null,
        createWorkbook: createBook,
        focusCell: (workbookName, sheetName, address) => {
          const target = openBook(workbookName, sheetName);
          focusCell(target, address);
          return { selected: address };
        },
      }),
    [
      engine,
      layout,
      active?.id,
      createBook,
      openBook,
      focusCell,
      patchPane,
      isLocked,
    ],
  );

  if (!ready)
    return (
      <div className="loading">
        <FileSpreadsheet size={34} />
        <p>Opening your workspace…</p>
      </div>
    );

  const renderPane = (pane: Pane) => {
    const sheets = engine.getOrderedSheetNames(pane.workbookName);
    const sheetName = sheets.includes(pane.sheetName)
      ? pane.sheetName
      : sheets[0];
    if (!sheetName)
      return <div className="empty-pane">No sheets in this workbook.</div>;
    const zoom = pane.zoom ?? 1;
    const sheetMeta = engine.getSheetMetadata({
      workbookName: pane.workbookName,
      sheetName,
    }) as Record<string, unknown> | undefined;
    const widths = sheetMeta?.arenaColumnWidths as
      Record<string, number> | undefined;
    const heights = sheetMeta?.arenaRowHeights as
      Record<string, number> | undefined;
    const selected = cellAddress(
      pane.workbookName,
      sheetName,
      pane.cell ?? "A1",
    );
    return (
      <section
        className={`workbook-pane ${pane.id === active?.id ? "active" : ""}`}
        aria-label={`${pane.workbookName} pane`}
        data-pane-id={pane.id}
        onPointerDownCapture={() => setActivePaneId(pane.id)}
        onFocusCapture={() => setActivePaneId(pane.id)}
      >
        <header className="pane-heading">
          <span className="pane-dot" />
          <FileSpreadsheet size={16} />
          <strong>{pane.workbookName}</strong>
          <span className="pane-sheet-count">
            {sheets.length} {sheets.length === 1 ? "sheet" : "sheets"}
          </span>
          <div className="pane-actions">
            <IconButton
              label={`Split ${pane.workbookName} right`}
              onClick={() => split(pane, "horizontal")}
            >
              <Columns2 size={15} />
            </IconButton>
            <IconButton
              label={`Split ${pane.workbookName} below`}
              onClick={() => split(pane, "vertical")}
            >
              <Rows2 size={15} />
            </IconButton>
            <IconButton
              label={
                maximized === pane.id
                  ? "Restore panes"
                  : `Maximize ${pane.workbookName}`
              }
              onClick={() =>
                setMaximized(maximized === pane.id ? null : pane.id)
              }
            >
              {maximized === pane.id ? (
                <Minimize2 size={15} />
              ) : (
                <Maximize2 size={15} />
              )}
            </IconButton>
            <IconButton
              label={`Close ${pane.workbookName} pane`}
              onClick={() => close(pane.id)}
            >
              <X size={15} />
            </IconButton>
          </div>
        </header>
        <div className="sheet-viewport">
          <div
            className="sheet-scale"
            style={{
              transform: zoom === 1 ? undefined : `scale(${zoom})`,
              width: `${100 / zoom}%`,
              height: `${100 / zoom}%`,
            }}
          >
            <WorkbookSheet
              theme={theme}
              key={`${pane.workbookName}:${sheetName}`}
              engine={engine}
              workbookName={pane.workbookName}
              sheetName={sheetName}
              viewId={pane.id}
              selectionManager={selection}
              clipboardManager={clipboard}
              aria-label={`${pane.workbookName}, ${sheetName}`}
              isSelected
              style={{ width: "100%", height: "100%" }}
              showFormulas={showFormulas}
              columnWidths={widths}
              rowHeights={heights}
              onCellDataChangeError={notifyError}
              selection={{
                initialState: {
                  selections: [
                    {
                      start: { row: selected.rowIndex, col: selected.colIndex },
                      end: {
                        row: { type: "number", value: selected.rowIndex },
                        col: { type: "number", value: selected.colIndex },
                      },
                    },
                  ],
                },
              }}
              onColumnWidthsChange={(value) =>
                engine.setSheetMetadata(
                  { workbookName: pane.workbookName, sheetName },
                  {
                    ...sheetMeta,
                    arenaColumnWidths: Object.fromEntries(
                      value instanceof Map ? value : Object.entries(value),
                    ),
                  },
                )
              }
              onRowHeightsChange={(value) =>
                engine.setSheetMetadata(
                  { workbookName: pane.workbookName, sheetName },
                  {
                    ...sheetMeta,
                    arenaRowHeights: Object.fromEntries(
                      value instanceof Map ? value : Object.entries(value),
                    ),
                  },
                )
              }
            />
          </div>
        </div>
        <footer className="sheet-footer">
          <div
            className="sheet-tabs"
            role="tablist"
            aria-label={`Sheets in ${pane.workbookName}`}
          >
            {sheets.map((name) => (
              <button
                key={name}
                role="tab"
                aria-selected={sheetName === name}
                onClick={() => {
                  patchPane(pane.id, { sheetName: name, cell: "A1" });
                  setActivePaneId(pane.id);
                }}
              >
                {name}
              </button>
            ))}
          </div>
          <IconButton
            label={`Add sheet to ${pane.workbookName}`}
            onClick={() => {
              const sheet = engine.createSheet({
                workbookName: pane.workbookName,
              });
              patchPane(pane.id, { sheetName: sheet.name, cell: "A1" });
            }}
          >
            <Plus size={16} />
          </IconButton>
          <label className="zoom-control">
            <span className="sr-only">Zoom {pane.workbookName}</span>
            <select
              value={zoom}
              onChange={(e) =>
                patchPane(pane.id, { zoom: Number(e.target.value) })
              }
            >
              {[0.5, 0.75, 1, 1.25, 1.5, 2].map((v) => (
                <option value={v} key={v}>
                  {Math.round(v * 100)}%
                </option>
              ))}
            </select>
          </label>
        </footer>
      </section>
    );
  };

  return (
    <div
      className="app-shell"
      onDragEnter={(event) => {
        if (event.dataTransfer.types.includes("Files")) {
          event.preventDefault();
          dragDepth.current++;
          setDragging(true);
        }
      }}
      onDragOver={(event) => {
        if (event.dataTransfer.types.includes("Files")) event.preventDefault();
      }}
      onDragLeave={() => {
        dragDepth.current = Math.max(0, dragDepth.current - 1);
        if (!dragDepth.current) setDragging(false);
      }}
      onDrop={(event) => {
        event.preventDefault();
        dragDepth.current = 0;
        setDragging(false);
        void importDropped([...event.dataTransfer.files]);
      }}
    >
      <a href="#workspace" className="skip-link">
        Skip to workspace
      </a>

      <div
        className="app-body"
        inert={saveInProgress}
        aria-busy={saveInProgress}
      >
        {sidebar && (
          <ResizableSidebar
            side="left"
            width={leftWidth}
            onWidthChange={setLeftWidth}
            aria-label="Workbook library"
          >
            <div className="sidebar">
              <div className="sidebar-title">
                <div className="brand">
                  <span className="brand-icon">
                    <FileSpreadsheet size={19} />
                  </span>
                  <h1>
                    spreadsheet<span>arena</span>
                  </h1>
                </div>
                <IconButton
                  label="Hide sidebar"
                  onClick={() => setSidebar(false)}
                >
                  <PanelLeftClose size={16} />
                </IconButton>
              </div>
              <Button
                className="new-workbook"
                onClick={() =>
                  setDialog({
                    title: "New workbook",
                    label: "Workbook name",
                    description: "",
                    initial: "Untitled workbook",
                    action: createBook,
                  })
                }
              >
                <Plus size={17} />
                New workbook
              </Button>
              <div className="library-actions">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => upload.current?.click()}
                  disabled={busy}
                >
                  <ArrowUpFromLine size={15} />
                  {busy ? "Importing…" : "Import"}
                </Button>
                <DropdownMenu>
                  <DropdownMenuTrigger asChild>
                    <Button variant="outline" size="sm">
                      <FolderOpen size={15} />
                      Workspace
                      <ChevronDown size={13} />
                    </Button>
                  </DropdownMenuTrigger>
                  <DropdownMenuContent align="start">
                    <DropdownMenuItem onSelect={downloadWorkspace}>
                      Download workspace backup
                    </DropdownMenuItem>
                    <DropdownMenuItem onSelect={() => restore.current?.click()}>
                      Restore workspace backup…
                    </DropdownMenuItem>
                    <DropdownMenuSeparator />
                    <DropdownMenuCheckboxItem
                      checked={debugMode}
                      onCheckedChange={setDebugMode}
                    >
                      Debug mode
                    </DropdownMenuCheckboxItem>
                  </DropdownMenuContent>
                </DropdownMenu>
              </div>
              <label className="search-field">
                <Search size={15} />
                <input
                  aria-label="Filter workbooks"
                  placeholder="Find a workbook…"
                  value={query}
                  onChange={(e) => setQuery(e.target.value)}
                />
                <span>⌕</span>
              </label>
              <div className="library-label">
                WORKBOOKS <span>{books.length}</span>
              </div>
              <nav aria-label="Workbooks" className="workbook-list">
                {books
                  .filter((name) =>
                    name.toLowerCase().includes(query.toLowerCase()),
                  )
                  .map((name, index) => {
                    const open = allPanes.some((p) => p.workbookName === name);
                    const current = active?.workbookName === name;
                    const sheets = engine.getOrderedSheetNames(name);
                    return (
                      <div
                        className={`library-book ${current ? "current" : ""}`}
                        key={name}
                      >
                        <button
                          className="book-button"
                          aria-label={`Open ${name}`}
                          aria-current={current ? "true" : undefined}
                          onClick={() => {
                            try {
                              openBook(name);
                            } catch (e) {
                              notifyError(e);
                            }
                          }}
                        >
                          <span className={`book-icon tone-${index % 3}`}>
                            <FileSpreadsheet size={17} />
                          </span>
                          <span className="book-name">
                            {name}
                            <small>
                              {sheets.length}{" "}
                              {sheets.length === 1 ? "sheet" : "sheets"}
                              {open ? " · open" : ""}
                            </small>
                          </span>
                          <ChevronDown size={13} />
                        </button>
                        {current && (
                          <div className="book-sheets">
                            {sheets.map((sheet) => (
                              <button
                                key={sheet}
                                onClick={() => openBook(name, sheet)}
                                className={
                                  active?.sheetName === sheet ? "selected" : ""
                                }
                              >
                                <span>▦</span>
                                {sheet}
                              </button>
                            ))}
                          </div>
                        )}
                        <div className="book-utilities">
                          <button
                            aria-label={`Download ${name}`}
                            onClick={() => {
                              try {
                                download(
                                  `${name}.workbook.json`,
                                  exportWorkbook(engine, name),
                                );
                              } catch (e) {
                                notifyError(e);
                              }
                            }}
                          >
                            <ArrowDownToLine size={12} />
                            Export
                          </button>
                          <button
                            aria-label={`Remove ${name}`}
                            onClick={() =>
                              setDialog({
                                title: `Remove ${name}?`,
                                description:
                                  "This removes the workbook from this device. Formulas in other workbooks may refer to it. Export a copy first; removal can be undone.",
                                danger: true,
                                action: () => engine.removeWorkbook(name),
                              })
                            }
                          >
                            <X size={12} />
                            Remove
                          </button>
                        </div>
                      </div>
                    );
                  })}
              </nav>
              <div className="sidebar-footer">
                <a
                  className="github-link"
                  href="https://github.com/ricsam/spreadsheet-arena"
                  target="_blank"
                  rel="noreferrer"
                >
                  GitHub ↗
                </a>
                <div className="sidebar-footer-actions">
                  <IconButton
                    label={
                      theme === "dark"
                        ? "Switch to light mode"
                        : "Switch to dark mode"
                    }
                    onClick={toggleTheme}
                  >
                    {theme === "dark" ? <Sun size={17} /> : <Moon size={17} />}
                  </IconButton>
                  <IconButton
                    label="Help and automation"
                    onClick={() => setHelp(!help)}
                  >
                    <HelpCircle size={17} />
                  </IconButton>
                </div>
              </div>
            </div>
          </ResizableSidebar>
        )}
        <main id="workspace" className="workspace" tabIndex={-1}>
          <div className="formula-bar">
            {!sidebar && (
              <IconButton
                label="Show sidebar"
                onClick={() => {
                  setSidebar(true);
                  if (window.matchMedia("(max-width: 900px)").matches)
                    setInspector(false);
                }}
              >
                <FolderOpen size={18} />
              </IconButton>
            )}
            <form
              className="cell-jump"
              onSubmit={(event) => {
                event.preventDefault();
                if (active)
                  try {
                    focusCell(active, cell);
                  } catch (e) {
                    notifyError(e);
                  }
              }}
            >
              <input
                aria-label="Cell address"
                value={cell}
                disabled={!active}
                onChange={(event) => setCell(event.target.value)}
              />
              <button aria-label="Go to cell" disabled={!active}>
                <ChevronDown size={13} />
              </button>
            </form>
            <span className="formula-symbol">ƒx</span>
            <form
              className="formula-form"
              onSubmit={(event) => {
                event.preventDefault();
                commitFormula();
              }}
            >
              <input
                ref={formulaInput}
                aria-label="Cell value or formula"

                value={formula}
                disabled={!active}
                onChange={(event) => {
                  setFormula(event.target.value);
                  setFormulaDirty(true);
                }}
                onKeyDown={(event) => {
                  if (event.key === "Escape") {
                    setFormula(raw);
                    setFormulaDirty(false);
                  }
                }}
              />
              <button
                aria-label="Apply cell value"
                disabled={!active || !formulaDirty}
              >
                <Check size={16} />
              </button>
            </form>
            <IconButton
              label={inspector ? "Hide inspector" : "Show inspector"}
              onClick={() => {
                setInspector(!inspector);
                if (
                  !inspector &&
                  window.matchMedia("(max-width: 900px)").matches
                )
                  setSidebar(false);
              }}
            >
              {inspector ? (
                <PanelRightClose size={17} />
              ) : (
                <PanelRightOpen size={17} />
              )}
            </IconButton>
          </div>
          <div
            className="command-bar"
            role="toolbar"
            aria-label="Workbook actions"
          >
            <div className="save-controls">
              <Button
                size="sm"
                onClick={() => void saveWorkspace()}
                disabled={saveInProgress || busy || !storageAvailable}
                title="Save workspace on this device (Ctrl/Cmd+S)"
              >
                <Save size={15} /> Save
              </Button>
              <select
                aria-label="Save contents"
                value={includeEvaluationState ? "warm" : "compact"}
                disabled={saveInProgress}
                onChange={(event) =>
                  setIncludeEvaluationState(event.target.value === "warm")
                }
              >
                <option value="compact">Workbook only (no cache)</option>
                <option value="warm">Workbook + calculation cache</option>
              </select>
            </div>
            <span className="toolbar-divider" />
            <div className="history-controls">
              <IconButton
                label="Undo"
                disabled={!engine.canUndo()}
                onClick={() => engine.undo()}
              >
                <Undo2 size={16} />
              </IconButton>
              <IconButton
                label="Redo"
                disabled={!engine.canRedo()}
                onClick={() => engine.redo()}
              >
                <Redo2 size={16} />
              </IconButton>
            </div>
            <span className="toolbar-divider" />
            <ClipboardActions
              engine={engine}
              selection={selection}
              clipboard={clipboard}
              active={active}
              isLocked={isLocked}
              onError={notifyError}
            />
            <span className="toolbar-divider" />
            <StyleActions
              engine={engine}
              selection={selection}
              active={active}
              revision={revision}
              isLocked={isLocked}
              onError={notifyError}
            />
            <span className="toolbar-divider" />
            <IconButton
              label="Show formulas"
              aria-pressed={showFormulas}
              onClick={() => setShowFormulas(!showFormulas)}
            >
              ƒx
            </IconButton>
            <div className="command-spacer" />
            <IconButton
              label="Split active pane right"
              disabled={!active}
              onClick={() => active && split(active, "horizontal")}
            >
              <Columns2 size={16} />
            </IconButton>
            <IconButton
              label="Split active pane below"
              disabled={!active}
              onClick={() => active && split(active, "vertical")}
            >
              <Rows2 size={16} />
            </IconButton>
          </div>
          {debugMode && (
            <section className="debug-panel" aria-label="Debug tools">
              <Bug size={15} />
              <Button
                variant="outline"
                size="sm"
                onClick={clearFormulaCache}
                disabled={!ready || busy || !books.length}
              >
                <RotateCcw size={15} />
                Clear formula cache
              </Button>
              <IconButton
                label="Disable debug mode"
                onClick={() => setDebugMode(false)}
              >
                <X size={16} />
              </IconButton>
            </section>
          )}
          {notice && (
            <div
              className={`notice ${notice.kind}`}
              role={notice.kind === "error" ? "alert" : "status"}
            >
              <span>{notice.message}</span>
              <IconButton
                label="Dismiss notification"
                onClick={() => setNotice(null)}
              >
                <X size={16} />
              </IconButton>
            </div>
          )}
          <div className="tile-stage">
            {layout ? (
              maximized && allPanes.some((p) => p.id === maximized) ? (
                renderPane(allPanes.find((p) => p.id === maximized)!)
              ) : (
                <Tiles
                  layout={layout}
                  renderPane={renderPane}
                  onResize={(id, ratio) =>
                    setLayout(
                      (current) => current && resizeSplit(current, id, ratio),
                    )
                  }
                />
              )
            ) : (
              <div className="empty-workspace">
                <div className="empty-illustration">
                  <FileSpreadsheet size={45} />
                </div>
                <button
                  className="button primary"
                  onClick={() => createBook("Untitled workbook")}
                >
                  <Plus size={16} />
                  Create a workbook
                </button>
                <button
                  className="text-button"
                  onClick={() => upload.current?.click()}
                >
                  Import a file instead
                </button>
              </div>
            )}
          </div>
          <footer className="workspace-status">
            <span>
              <span className="status-dot" />
              {busy ? "Importing" : "Ready"}
              <span className="status-separator">|</span>
              {active
                ? `${active.workbookName} / ${active.sheetName}`
                : "No workbook open"}
            </span>
            {!sidebar && (
              <div className="collapsed-sidebar-actions">
                <IconButton
                  label={
                    theme === "dark"
                      ? "Switch to light mode"
                      : "Switch to dark mode"
                  }
                  onClick={toggleTheme}
                >
                  {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}
                </IconButton>
                <IconButton
                  label="Help and automation"
                  onClick={() => setHelp(!help)}
                >
                  <HelpCircle size={16} />
                </IconButton>
              </div>
            )}
            <span role="status">
              {active?.cell ?? "A1"}
              <span className="status-separator">|</span>
              {saving}
            </span>
          </footer>
        </main>
        {inspector && (
          <ResizableSidebar
            side="right"
            width={rightWidth}
            onWidthChange={setRightWidth}
            minWidth={260}
            maxWidth={440}
            aria-label="Workbook inspector"
          >
            <div className="inspector-heading">
              <span
                className="inspector-context"
                title={
                  active
                    ? `${active.workbookName} / ${active.sheetName}`
                    : undefined
                }
              >
                {active?.workbookName ?? "Workbook"}
              </span>
              <IconButton
                label="Close inspector"
                onClick={() => setInspector(false)}
              >
                <X size={17} />
              </IconButton>
            </div>
            <WorkbookInspector
              engine={engine}
              selection={selection}
              active={active}
              revision={revision}
              isLocked={isLocked}
              onError={notifyError}
            />
          </ResizableSidebar>
        )}
      </div>
      <input
        ref={upload}
        type="file"
        className="sr-only"
        aria-label="Import workbook files"
        disabled={saveInProgress || busy}
        multiple
        accept=".xlsx,.xlsm,.csv,.tsv,.json"
        onChange={(event) => {
          void importDropped([...(event.target.files ?? [])]);
          event.target.value = "";
        }}
      />
      <input
        ref={restore}
        type="file"
        className="sr-only"
        aria-label="Restore workspace file"
        disabled={saveInProgress || busy}
        accept=".json"
        onChange={async (event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (!file || saveLocked.current || importInFlight.current) return;
          importInFlight.current = true;
          setBusy(true);
          try {
            if (file.size > MAX_WORKSPACE_FILE_BYTES)
              throw new Error(
                "Workspace backups must be no larger than 256 MiB.",
              );
            const { snapshot, engine: restoredEngine } = restoreSnapshot(
              JSON.parse(await file.text()),
            );
            setDialog({
              title: "Restore workspace?",
              description:
                "This replaces the open workbooks and pane layout, discarding unsaved changes. Download a backup first if needed. Use Save afterwards to replace the saved workspace on this device.",
              action: () => {
                if (saveLocked.current) throw new Error("Saving in progress.");
                // Replacement is explicit and user-confirmed; subsequent saves use CAS.
                setWorkspaceWriter(createWorkspaceWriter(null, true));
                setEngine(restoredEngine);
                setLayout(snapshot.layout);
                setActivePaneId(snapshot.activePaneId);
                setMaximized(snapshot.view?.maximized ?? null);
                setSidebar(snapshot.view?.sidebar ?? true);
                setShowFormulas(snapshot.view?.showFormulas ?? false);
                setSavedState(null);
                setSaveError(false);
                setStorageAvailable(true);
              },
            });
          } catch (e) {
            notifyError(e);
          } finally {
            importInFlight.current = false;
            setBusy(false);
          }
        }}
      />
      {saveInProgress && (
        <dialog
          ref={saveDialog}
          aria-labelledby="save-title"
          onCancel={(event) => event.preventDefault()}
        >
          <h2 id="save-title">Saving workspace…</h2>
          <p role="status">
            Editing is temporarily disabled. Keep this tab open until saving
            finishes.
          </p>
          <p>
            {includeEvaluationState
              ? "Including warm calculation cache."
              : "Saving workbook data without calculation cache."}{" "}
            Large workspaces may briefly pause the browser.
          </p>
        </dialog>
      )}
      {dragging && (
        <div className="drop-overlay">
          <ArrowUpFromLine size={42} />
          <h2>Drop into your workspace</h2>
          <p>Excel, CSV, TSV, or formula-engine JSON</p>
        </div>
      )}
      {dialog && (
        <PromptDialog
          key={dialog.title}
          state={dialog}
          onClose={() => setDialog(null)}
        />
      )}
      {help && (
        <div
          className="help-panel"
          role="region"
          aria-label="Help and automation"
        >
          <div className="dialog-heading">
            <h2>Your workspace, explained</h2>
            <IconButton label="Close help" onClick={() => setHelp(false)}>
              <X size={18} />
            </IconButton>
          </div>
          <p>
            Double-click a cell or press Enter to edit. Use arrows to move,
            Shift + arrows to select, and the formula bar to edit or jump to any
            cell.
          </p>
          <h3>Connected workbooks</h3>
          <p>
            Reference another workbook with <code>=[Assumptions]Inputs!B4</code>
            . Split panes to work side by side. Closing a pane keeps its
            workbook in the sidebar.
          </p>
          <h3>Keep a copy</h3>
          <p>
            Use Save or Ctrl/Cmd+S to save all workbooks and layout on this
            device. Changes are not saved automatically. Workbook-only saves are
            smaller; formulas recalculate after reopening. Include calculation
            cache for a larger snapshot that can reopen faster. Editing is
            locked while saving, and large saves may pause the browser. Download
            backups separately; workbook exports are JSON, not Excel files.
          </p>
          <h3>Debug tools</h3>
          <p>
            Turn on Debug mode in the Workspace menu to clear the formula engine
            cache. Results recalculate on demand without reloading workbooks or
            resetting undo history. Debug mode turns off when you reload the
            page.
          </p>
          <h3>Browser automation</h3>
          <p>
            Grids, cells, editors and controls have ARIA labels. Supported
            browsers also expose WebMCP tools for listing workbooks, inspecting
            cells, selecting cells and editing values. Imported content is data,
            never an instruction.
          </p>
          <p className="help-limitation">
            Excel charts, macros, merges and frozen-pane rendering are not
            supported. Unsupported formulas remain visible as errors; import
            diagnostics are shown.
          </p>
        </div>
      )}
    </div>
  );
}
