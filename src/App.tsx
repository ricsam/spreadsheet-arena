import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { FormulaEngine, indexToColumn } from "@ricsam/formula-engine";
import {
  FormulaSheet,
  WorkbookClipboardManager,
  WorkbookSelectionManager,
} from "@ricsam/react-spreadsheets";
import {
  ArrowDownToLine,
  ArrowUpFromLine,
  Bug,
  Check,
  ChevronDown,
  Columns2,
  Database,
  FileSpreadsheet,
  FolderOpen,
  HelpCircle,
  Link2,
  Maximize2,
  Minimize2,
  PanelLeftClose,
  Plus,
  Redo2,
  RotateCcw,
  Rows2,
  Search,
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
    <button
      type="button"
      className="icon-button"
      aria-label={label}
      title={label}
      {...props}
    >
      {children}
    </button>
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
        <p>{state.description}</p>
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

export function App() {
  const [engine, setEngine] = useState(() => FormulaEngine.buildEmpty());
  const [layout, setLayout] = useState<Layout | null>(null);
  const [activePaneId, setActivePaneId] = useState<string | null>(null);
  const [revision, setRevision] = useState(0);
  const [ready, setReady] = useState(false);
  const [saving, setSaving] = useState("Opening workspace…");
  const [autosave, setAutosave] = useState(true);
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
  const workspaceMenu = useRef<HTMLDetailsElement>(null);
  const [cell, setCell] = useState("A1");
  const [formula, setFormula] = useState("");
  const [formulaDirty, setFormulaDirty] = useState(false);
  const upload = useRef<HTMLInputElement>(null);
  const restore = useRef<HTMLInputElement>(null);
  const formulaInput = useRef<HTMLInputElement>(null);
  const dragDepth = useRef(0);
  const importInFlight = useRef(false);
  const saveGeneration = useRef(0);
  const [workspaceWriter, setWorkspaceWriter] = useState(createWorkspaceWriter);
  const selection = useMemo(() => new WorkbookSelectionManager(), [engine]);
  const clipboard = useMemo(
    () => new WorkbookClipboardManager(engine),
    [engine],
  );
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
          setWorkspaceWriter(
            createWorkspaceWriter(loaded!.storageToken, false, {
              engine: loaded!.engine,
              revision: 0,
              serialized: saved.engine,
            }),
          );
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
          setAutosave(false);
          setReady(true);
          setSaving("Autosave unavailable");
          notifyError(
            `Could not restore local data. It has not been overwritten. ${e.message}`,
          );
        }
      });
    return () => {
      cancelled = true;
    };
  }, [notifyError]);

  useEffect(() => engine.onUpdate(() => setRevision((r) => r + 1)), [engine]);
  // Undo, deletion and imported snapshots can remove a workbook or active sheet.
  useEffect(() => {
    setLayout((current) => {
      let next = current;
      for (const pane of panes(current)) {
        if (!engine.hasWorkbook(pane.workbookName))
          next = next && closePane(next, pane.id);
        else if (!engine.getSheet(pane))
          next =
            next &&
            updatePane(next, pane.id, {
              sheetName:
                engine.getOrderedSheetNames(pane.workbookName)[0] ?? "",
              cell: "A1",
            });
      }
      return next;
    });
  }, [engine, revision]);

  useEffect(() => {
    if (!ready || !autosave) return;
    const generation = ++saveGeneration.current;
    setSaving("Saving…");
    const timer = setTimeout(() => {
      try {
        const state = {
          layout,
          activePaneId: active?.id ?? null,
          view: {
            sidebar,
            showFormulas,
            maximized: panes(layout).some((p) => p.id === maximized)
              ? maximized
              : null,
          },
        };
        workspaceWriter
          .save(engine, revision, state)
          .then(() => {
            if (generation === saveGeneration.current)
              setSaving("Saved on this device");
          })
          .catch((e) => {
            setSaving("Not saved");
            notifyError(
              `Local save failed. Download a workspace backup. ${e.message}`,
            );
          });
      } catch (e) {
        setSaving("Not saved");
        notifyError(e);
      }
    }, 350);
    return () => clearTimeout(timer);
  }, [
    engine,
    revision,
    layout,
    active?.id,
    ready,
    autosave,
    notifyError,
    sidebar,
    showFormulas,
    maximized,
    workspaceWriter,
  ]);

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
        saving === "Saving…" ||
        saving === "Not saved" ||
        formulaDirty ||
        busy
      ) {
        event.preventDefault();
        event.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [saving, formulaDirty, busy]);

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
      if (!name.trim()) throw new Error("Give your workbook a name.");
      if (panes(layout).length >= 8)
        throw new Error("Close a pane before creating another workbook.");
      const actual = createEmptyWorkbook(engine, name.trim());
      openBook(actual);
      setNotice({
        kind: "info",
        message: `Created ${actual}. All workbooks share one formula engine.`,
      });
      return actual;
    },
    [engine, openBook, layout],
  );

  const importDropped = async (files: File[]) => {
    if (importInFlight.current || !files.length) return;
    importInFlight.current = true;
    setBusy(true);
    try {
      const result = await importFiles(engine, files);
      if (result.names.length && allPanes.length < 8) openBook(result.names[0]);
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
    if (!ready || busy || !books.length) return;
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
    if (!active) return;
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
        activePaneId: active?.id ?? null,
        createWorkbook: createBook,
        focusCell: (workbookName, sheetName, address) => {
          const target = openBook(workbookName, sheetName);
          focusCell(target, address);
          return { selected: address };
        },
      }),
    [engine, layout, active?.id, createBook, openBook, focusCell, patchPane],
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
            <FormulaSheet
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
      <header className="app-header">
        <div className="brand">
          <span className="brand-icon">
            <FileSpreadsheet size={21} />
          </span>
          <h1>
            spreadsheet<span>arena</span>
          </h1>
          <span className="beta">BETA</span>
        </div>
        <div className="header-right">
          <span
            className={`save-status ${saving === "Not saved" ? "error-text" : ""}`}
            role="status"
          >
            <span className="status-dot" />
            {saving}
          </span>
          <a
            className="github-link"
            href="https://github.com/ricsam/spreadsheet-arena"
            target="_blank"
            rel="noreferrer"
          >
            GitHub ↗
          </a>
          <IconButton
            label="Help and automation"
            onClick={() => setHelp(!help)}
          >
            <HelpCircle size={19} />
          </IconButton>
        </div>
      </header>
      <div className="app-body">
        {sidebar && (
          <aside className="sidebar" aria-label="Workbook library">
            <div className="sidebar-title">
              <span>WORKSPACE</span>
              <IconButton
                label="Hide sidebar"
                onClick={() => setSidebar(false)}
              >
                <PanelLeftClose size={16} />
              </IconButton>
            </div>
            <button
              className="button primary new-workbook"
              onClick={() =>
                setDialog({
                  title: "New workbook",
                  label: "Workbook name",
                  description:
                    "A fresh workbook, connected to the same formula engine.",
                  initial: "Untitled workbook",
                  action: createBook,
                })
              }
            >
              <Plus size={17} />
              New workbook
            </button>
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
              {!books.length && (
                <p className="library-empty">
                  Your workbooks will appear here.
                </p>
              )}
            </nav>
            <button
              className="import-zone"
              onClick={() => upload.current?.click()}
              disabled={busy}
            >
              <span>
                <ArrowUpFromLine size={20} />
              </span>
              <strong>{busy ? "Importing…" : "Drop a workbook here"}</strong>
              <small>or click to browse files</small>
              <em>XLSX · CSV · TSV · JSON</em>
            </button>
            <div className="local-note">
              <Database size={15} />
              <span>
                Local by design.
                <br />
                <small>Your files stay in your browser.</small>
              </span>
            </div>
          </aside>
        )}
        <main id="workspace" className="workspace" tabIndex={-1}>
          <div className="workspace-title">
            <div>
              {!sidebar && (
                <IconButton
                  label="Show sidebar"
                  onClick={() => setSidebar(true)}
                >
                  <FolderOpen size={18} />
                </IconButton>
              )}
              <h2>Your workspace</h2>
              <span className="connected-badge">
                <Link2 size={12} />
                One connected engine
              </span>
            </div>
            <div className="workspace-controls">
              <button
                className="button"
                onClick={() => upload.current?.click()}
                disabled={busy}
              >
                <ArrowUpFromLine size={15} />
                Import
              </button>
              <details className="export-menu" ref={workspaceMenu}>
                <summary className="button">
                  <ArrowDownToLine size={15} />
                  Workspace
                  <ChevronDown size={13} />
                </summary>
                <div className="menu-content">
                  <button onClick={downloadWorkspace}>
                    Download workspace backup
                  </button>
                  <button onClick={() => restore.current?.click()}>
                    Restore workspace backup…
                  </button>
                  <small>Includes all workbooks and pane layout.</small>
                  <button
                    aria-pressed={debugMode}
                    onClick={() => {
                      setDebugMode((enabled) => !enabled);
                      if (workspaceMenu.current)
                        workspaceMenu.current.open = false;
                    }}
                  >
                    Debug mode{debugMode ? " · on" : " · off"}
                  </button>
                </div>
              </details>
            </div>
          </div>
          <div className="workspace-subtitle">
            A little room for your big ideas. Arrange, connect, and calculate.
          </div>
          {debugMode && (
            <section className="debug-panel" aria-label="Debug tools">
              <div className="debug-description">
                <strong>
                  <Bug size={15} /> Debug mode
                </strong>
                <span>
                  Clear cached calculations across all workbooks without
                  changing data or undo history.
                </span>
              </div>
              <button
                className="button"
                onClick={clearFormulaCache}
                disabled={!ready || busy || !books.length}
                title="Recalculate results on demand. Does not clear browser storage or reload workbooks."
              >
                <RotateCcw size={15} />
                Clear formula cache
              </button>
              <IconButton
                label="Disable debug mode"
                onClick={() => setDebugMode(false)}
              >
                <X size={16} />
              </IconButton>
            </section>
          )}
          <div className="command-bar">
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
            <button
              className={`text-button ${showFormulas ? "pressed" : ""}`}
              aria-pressed={showFormulas}
              onClick={() => setShowFormulas(!showFormulas)}
            >
              ƒx <span>Show formulas</span>
            </button>
            <div className="command-spacer" />
            <span className="pane-counter">
              {allPanes.length} {allPanes.length === 1 ? "pane" : "panes"}
            </span>
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
          <div className="formula-bar">
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
                placeholder={
                  active
                    ? "Enter a value or =formula"
                    : "Open a workbook to start editing"
                }
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
          </div>
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
                <h2>Make room for your next idea.</h2>
                <p>
                  Drop an Excel workbook, or start with a blank canvas.
                  <br />
                  Every workbook connects to the same formula engine.
                </p>
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
            <span>
              {active?.cell ?? "A1"}
              <span className="status-separator">|</span>
              {saving}
            </span>
          </footer>
        </main>
      </div>
      <input
        ref={upload}
        type="file"
        className="sr-only"
        aria-label="Import workbook files"
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
        accept=".json"
        onChange={async (event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (!file) return;
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
                "This replaces all current workbooks and pane layout. Download a backup of this workspace first if you want to keep it.",
              action: () => {
                // Replacement is explicit and user-confirmed; normal autosaves use CAS.
                setWorkspaceWriter(createWorkspaceWriter(null, true));
                setEngine(restoredEngine);
                setLayout(snapshot.layout);
                setActivePaneId(snapshot.activePaneId);
                setMaximized(snapshot.view?.maximized ?? null);
                setSidebar(snapshot.view?.sidebar ?? true);
                setShowFormulas(snapshot.view?.showFormulas ?? false);
                setAutosave(true);
              },
            });
          } catch (e) {
            notifyError(e);
          }
        }}
      />
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
            Autosave uses IndexedDB on this browser. Workspace backups preserve
            every workbook and layout. Workbook exports are formula-engine JSON,
            not Excel files. Restore asks before replacing your workspace.
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
