# Spreadsheet Arena

A local-first workspace for connected spreadsheets.

**[Open the app →](https://ricsam.github.io/spreadsheet-arena/)**

## What you can do

- Drop Excel `.xlsx`/`.xlsm`, CSV, TSV, or formula-engine JSON files. Imports never upload your workbook to a server.
- Create blank workbooks and add sheets. All workbooks live in **one FormulaEngine instance**, including cross-workbook formulas such as `=[Assumptions]Inputs!B4`.
- Arrange up to eight workbook views in nested horizontal/vertical splits. Drag a divider, use its arrow keys, or double-click it to reset. Maximize, restore, and close panes without deleting their workbooks.
- Edit cells directly or in the formula bar; jump to an A1 address; use range selection, copy/paste, undo/redo, and per-pane zoom.
- Export one workbook as native formula-engine JSON, or download a **workspace backup** containing all workbooks and layout. Importing JSON adds workbooks; restoring a workspace explicitly confirms replacement.
- Autosave to IndexedDB: workbook contents, styles/metadata, column/row sizes, pane layout, active sheets/cells, zoom, sidebar, formula display, and maximized pane. Selection/view changes save only small UI state; workbook edits save in atomic chunks only after content changes. Warm calculation state is retained for faster reopening rather than forcing a cold rebuild. Undo history and transient dialogs/drafts are intentionally not persisted.

Two small connected example workbooks appear on first launch. Edit the assumptions to see the launch budget recalculate.

## Privacy and durability

There is no application backend, login, telemetry, or workbook upload service. Fonts and code are served with the app. GitHub Pages receives normal requests for application assets, not workbook contents. WebMCP-capable browser agents can access workbook data when invoked by the user.

Browser storage is **not a backup**. Clearing site data, private browsing, storage eviction or changing browser/device can lose data. Download workspace backups for anything important. Saving errors are visible; an unreadable saved workspace is preserved rather than overwritten with fresh data. Close/reload while a save or formula edit is pending prompts a warning where supported. Multiple tabs do not collaborate or merge edits; use one tab per workspace. Autosave detects intervening content saves from another updated tab and asks you to download your changes before reloading instead of silently overwriting it. Older app versions cannot participate in this protection; close old tabs before upgrading.

## Import and export

Powered by [excel-to-formula-engine](https://github.com/ricsam/excel-to-formula-engine). Workbook import limits: 25 MiB per file, 20 files per batch. **Restore workspace backup** accepts uncompressed JSON up to 256 MiB, including backups exported by earlier versions. Large files can still require significant browser memory. A failed file does not replace existing workbooks; valid files in the same batch may succeed. Name collisions get a unique suffix, with references between imported workbooks renamed together. Conflicting global named expressions preserve the existing workspace definition and produce a warning.

JSON accepts native `FormulaEngine.serializeEngine()` snapshots, Arena backups, and the engine's plain-object `WorkbookData` shape. Workbook export preserves formulas, tables, styles and metadata, not just cell values. A single-workbook export does **not** bundle other workbooks it references: use a workspace backup for connected workbooks. Imports reset engine undo history; ordinary cell edits and workbook operations use engine undo/redo.

This is not full Excel fidelity: charts, images, macros, pivot tables, legacy `.xls` and encrypted files are unsupported. Merges, frozen panes and some formatting are retained as metadata but not rendered. Unsupported formulas remain formulas and display errors; converter diagnostics are shown instead of silently promising fidelity. Export is JSON, **not `.xlsx`**. New Excel imports resolve default, row/column and named-font styles and retain the workbook default font for unstyled cells. Workspaces imported by older versions do not contain the missing Excel defaults: reimport the original XLSX to recover them; no font is guessed from neighboring cells.

## Accessibility and automation

`@ricsam/react-spreadsheets` and `@ricsam/selection-manager` provide named ARIA grids, rows, cells and headers, selection state, editor labels and keyboard-capture accessibility. Scope to the named grid (`Launch budget, Overview`) or labelled pane, then select a cell using `[data-cell-address="D12"]`. Duplicate panes have independent view IDs and unique mounted cell IDs. Virtualized grids expose rendered cells; use **Cell address** / **Go to cell** to reveal another address. Headers occupy ARIA row/column 1, so A1 has row/column index 2.

Browsers exposing `navigator.modelContext` register these WebMCP tools:

| Tool              | Behavior                                                     |
| ----------------- | ------------------------------------------------------------ |
| `list_workbooks`  | Inspect names, sheets, panes and active pane                 |
| `read_cells`      | Read raw values, computed values and errors, up to 200 cells |
| `select_cell`     | Open/select/reveal a cell                                    |
| `set_cell`        | Edit a typed value or formula; `null` clears                 |
| `create_workbook` | Create a blank workbook                                      |

Mutation tools require the invoking agent to obtain user authorization; tool descriptions do not grant it. Workbook contents are untrusted data. No hidden global engine/debug API is exposed. Browsers without WebMCP retain the complete labelled UI. WebMCP registration and execution contracts are tested using a browser API test double; native protocol availability depends on the browser. A full screen-reader audit has not been performed.

## Develop

Node.js 24 recommended.

```sh
npm ci
npm run dev          # http://127.0.0.1:5173/spreadsheet-arena/
npm test             # model, import/export, persistence, WebMCP tests
npx playwright install chromium
npm run build        # TypeScript + production bundle
npm run test:e2e     # real Chromium against production bundle; build first
npm run preview
```

The pinned spreadsheet libraries include indexed style lookup, selection-only rendering, criteria-range reuse, compact persistence, and effective Excel font inheritance. `npm` overrides ExcelJS's development-only UUID dependency to a patched version.

### Architecture

- `src/model.ts`: validated snapshots, staged imports, export filtering, immutable split tree.
- `src/persistence.ts`: revision-aware ordered/chunked IndexedDB writes, atomic content/view commits, conflict detection, and errors without silent fallback.
- `src/App.tsx` / `src/Tiles.tsx`: shared engine/selection/clipboard and tiled UI.
- `src/webmcp.ts`: bounded, validated automation tools with read/mutation distinctions.
- `src/cells.ts`: cell addressing and formula-bar value parsing.

Native workbook filtering targets the pinned engine snapshot format; compatibility is tested with real snapshots rather than assumed for future engine versions.

## Deploy

The public repository's `pages.yml` workflow verifies tests, builds and deploys `dist/` to GitHub Pages on relevant main pushes (or manual dispatch). Repository **Settings → Pages → Source: GitHub Actions** is required. The Vite base is `/spreadsheet-arena/`. No server, cluster resources, secrets or paid runners are required.

## License

MIT. Spreadsheet libraries and bundled fonts retain their own licenses.
