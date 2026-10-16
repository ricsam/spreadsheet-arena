import "fake-indexeddb/auto";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { deleteDB, openDB } from "idb";
import { FormulaEngine } from "@ricsam/formula-engine";
import { createDemoEngine, makeSnapshot, type WorkspaceView } from "./model";
import {
  readWorkspace,
  writeWorkspace,
  loadWorkspace,
  createWorkspaceWriter,
  ENGINE_CHUNK_CHARS,
} from "./persistence";

beforeEach(async () => {
  await deleteDB("spreadsheet-arena");
});
const snapshot = () => makeSnapshot(createDemoEngine(), null, null);
const warm = { includeEvaluationState: true };
const compact = { includeEvaluationState: false };
const view = { layout: null, activePaneId: null };

function formulaEngine() {
  const engine = new FormulaEngine();
  engine.addWorkbook({ workbookName: "Book" });
  engine.addSheet({ workbookName: "Book", sheetName: "Sheet" });
  engine.setCellContent(
    { workbookName: "Book", sheetName: "Sheet", rowIndex: 0, colIndex: 0 },
    21,
  );
  engine.setCellContent(
    { workbookName: "Book", sheetName: "Sheet", rowIndex: 0, colIndex: 1 },
    "=A1*2",
  );
  return engine;
}
const formulaCell = {
  workbookName: "Book",
  sheetName: "Sheet",
  rowIndex: 0,
  colIndex: 1,
};

describe("IndexedDB persistence", () => {
  it("returns null only when no saved workspace exists", async () => {
    expect(await readWorkspace()).toBeNull();
  });
  it("saves and restores the native engine snapshot", async () => {
    const value = snapshot();
    await writeWorkspace(value);
    expect(await readWorkspace()).toEqual(value);
  });
  it.each([false, true])(
    "roundtrips explicit saves with includeEvaluationState=%s",
    async (includeEvaluationState) => {
      const engine = formulaEngine();
      expect(engine.getCellValue(formulaCell)).toBe(42);
      const options = { includeEvaluationState };
      const expected = engine.serializeEngine(options);
      const serialize = vi.spyOn(engine, "serializeEngine");
      await createWorkspaceWriter().save(engine, view, options);
      expect(serialize).toHaveBeenCalledExactlyOnceWith(options);
      const loaded = (await loadWorkspace())!;
      expect(loaded.snapshot.engine).toBe(expected);
      expect(loaded.engine.getCellValue(formulaCell)).toBe(42);
    },
  );
  it("changes mode on unchanged data and freshly serializes repeated explicit saves", async () => {
    const engine = formulaEngine();
    engine.getCellValue(formulaCell);
    const writer = createWorkspaceWriter();
    const serialize = vi.spyOn(engine, "serializeEngine");
    const ids = [];
    for (const options of [warm, compact, warm, warm]) {
      await writer.save(engine, view, options);
      const loaded = (await loadWorkspace())!;
      expect(loaded.snapshot.engine).toBe(serialize.mock.results.at(-1)!.value);
      ids.push(loaded.storageToken);
    }
    expect(serialize.mock.calls).toEqual([[warm], [compact], [warm], [warm]]);
    expect(serialize.mock.results[0].value).not.toBe(
      serialize.mock.results[1].value,
    );
    expect(new Set(ids).size).toBe(4);
  });
  it("recaptures newly warmed caches without a content mutation", async () => {
    const engine = formulaEngine();
    const writer = createWorkspaceWriter();
    await writer.save(engine, view, warm);
    const cold = (await readWorkspace())!.engine;
    expect(engine.getCellValue(formulaCell)).toBe(42);
    const expected = engine.serializeEngine(warm);
    expect(expected).not.toBe(cold);
    await writer.save(engine, view, warm);
    expect((await readWorkspace())!.engine).toBe(expected);
  });
  it("captures queued saves synchronously in order and detaches mutable views", async () => {
    const engine = createDemoEngine();
    const writer = createWorkspaceWriter();
    const state: WorkspaceView = {
      layout: {
        kind: "pane",
        pane: {
          id: "p",
          workbookName: "Launch budget",
          sheetName: "Overview",
          cell: "A1",
        },
      },
      activePaneId: "p",
    };
    const serialize = vi.spyOn(engine, "serializeEngine");
    const first = writer.save(engine, state, compact);
    if (state.layout?.kind === "pane") state.layout.pane.cell = "B2";
    engine.setWorkbookMetadata("Launch budget", { saved: true });
    const second = writer.save(engine, state, warm);
    const expected = structuredClone(state);
    if (state.layout?.kind === "pane") state.layout.pane.cell = "C3";
    engine.setWorkbookMetadata("Launch budget", { notSaved: true });
    expect(serialize).toHaveBeenCalledTimes(2);
    await Promise.all([first, second]);
    const loaded = (await loadWorkspace())!;
    expect(loaded.snapshot.layout).toEqual(expected.layout);
    expect(loaded.engine.getWorkbookMetadata("Launch budget")).toEqual({
      saved: true,
    });
  });
  it("does not erase corrupted stored data on read or an unconfirmed fresh save", async () => {
    await readWorkspace();
    const db = await openDB("spreadsheet-arena", 1);
    const broken = {
      format: "spreadsheet-arena",
      version: 99,
      engine: "broken",
    };
    await db.put("workspace", broken, "current");
    await expect(readWorkspace()).rejects.toThrow(
      "original data has been kept",
    );
    await expect(
      createWorkspaceWriter().save(createDemoEngine(), view, compact),
    ).rejects.toThrow("Another tab changed");
    expect(await db.get("workspace", "current")).toEqual(broken);
    db.close();
  });
  it("rejects invalid saves without replacing an existing valid save", async () => {
    const value = snapshot();
    await writeWorkspace(value);
    await expect(writeWorkspace({ ...value, engine: "{}" })).rejects.toThrow();
    expect(await readWorkspace()).toEqual(value);
    await writeWorkspace(value);
  });
  it.each([false, true])(
    "loads existing data without writing or serializing (legacy=%s)",
    async (legacy) => {
      const value = snapshot();
      await readWorkspace();
      const db = await openDB("spreadsheet-arena", 1);
      if (legacy) await db.put("workspace", value, "current");
      else await writeWorkspace(value);
      const before = await db.getAll("workspace");
      const serialize = vi.spyOn(FormulaEngine.prototype, "serializeEngine");
      const reset = vi.spyOn(
        FormulaEngine.prototype,
        "resetToSerializedEngine",
      );
      const loaded = (await loadWorkspace())!;
      expect(reset).toHaveBeenCalledTimes(1);
      expect(serialize).not.toHaveBeenCalled();
      reset.mockRestore();
      serialize.mockRestore();
      expect(loaded.snapshot).toEqual(value);
      expect(await db.getAll("workspace")).toEqual(before);
      // Migration happens only on this explicit save, in the selected mode.
      const expected = loaded.engine.serializeEngine(compact);
      await createWorkspaceWriter(loaded.storageToken).save(
        loaded.engine,
        loaded.snapshot,
        compact,
      );
      expect((await db.get("workspace", "current")).format).toBe(
        "spreadsheet-arena-storage",
      );
      expect((await readWorkspace())!.engine).toBe(expected);
      db.close();
    },
  );
  it("chunks large content and atomically preserves all old data on failure, then retries freshly", async () => {
    const engine = createDemoEngine();
    engine.setWorkbookMetadata("Launch budget", {
      payload: "x".repeat(ENGINE_CHUNK_CHARS * 2),
    });
    const writer = createWorkspaceWriter();
    await writer.save(engine, view, compact);
    const db = await openDB("spreadsheet-arena", 1);
    const before = await db.get("workspace", "current");
    const keys = await db.getAllKeys("workspace");
    const values = await db.getAll("workspace");
    expect(before.chunks).toBeGreaterThan(2);
    for (const key of keys) {
      if (String(key).startsWith("engine:"))
        expect((await db.get("workspace", key)).length).toBeLessThanOrEqual(
          ENGINE_CHUNK_CHARS,
        );
    }
    const originalPut = IDBObjectStore.prototype.put;
    const fail = vi
      .spyOn(IDBObjectStore.prototype, "put")
      .mockImplementation(function (this: IDBObjectStore, value, key) {
        if (key === "current")
          throw new DOMException("full", "QuotaExceededError");
        return originalPut.call(this, value, key);
      });
    engine.setWorkbookMetadata("Launch budget", { payload: "changed" });
    try {
      await expect(writer.save(engine, view, warm)).rejects.toThrow("full");
    } finally {
      fail.mockRestore();
    }
    expect(await db.getAllKeys("workspace")).toEqual(keys);
    expect(await db.getAll("workspace")).toEqual(values);
    engine.setWorkbookMetadata("Launch budget", { payload: "retry" });
    await writer.save(engine, view, compact);
    expect(
      (await loadWorkspace())!.engine.getWorkbookMetadata("Launch budget"),
    ).toEqual({ payload: "retry" });
    expect(
      (await db.getAllKeys("workspace")).some((k) =>
        String(k).includes(before.engineId),
      ),
    ).toBe(false);
    db.close();
  });
  it("retries after serialization failure without poisoning the queue", async () => {
    const engine = createDemoEngine();
    const writer = createWorkspaceWriter();
    const serialize = vi
      .spyOn(engine, "serializeEngine")
      .mockImplementationOnce(() => {
        throw new Error("capture failed");
      });
    await expect(writer.save(engine, view, warm)).rejects.toThrow(
      "capture failed",
    );
    expect(await readWorkspace()).toBeNull();
    await writer.save(engine, view, compact);
    expect(serialize).toHaveBeenCalledTimes(2);
    expect(await loadWorkspace()).not.toBeNull();
  });
  it.each([false, true])(
    "guards first explicit save after load against intervening writes (legacy=%s)",
    async (legacy) => {
      const initial = snapshot();
      await readWorkspace();
      if (legacy) {
        const db = await openDB("spreadsheet-arena", 1);
        await db.put("workspace", initial, "current");
        db.close();
      } else await writeWorkspace(initial);
      const loaded = (await loadWorkspace())!;
      const writer = createWorkspaceWriter(loaded.storageToken);
      const newer = createDemoEngine();
      newer.setWorkbookMetadata("Launch budget", { newer: true });
      await writeWorkspace(makeSnapshot(newer, null, null));
      for (const options of [warm, compact])
        await expect(
          writer.save(loaded.engine, loaded.snapshot, options),
        ).rejects.toThrow("Another tab changed");
      expect(
        (await loadWorkspace())!.engine.getWorkbookMetadata("Launch budget"),
      ).toEqual({ newer: true });
    },
  );
  it("only allows one fresh writer to claim an initially empty workspace", async () => {
    const first = createWorkspaceWriter(),
      second = createWorkspaceWriter();
    const results = await Promise.allSettled([
      first.save(createDemoEngine(), view, compact),
      second.save(createDemoEngine(), view, warm),
    ]);
    expect(results.map((r) => r.status)).toEqual(["fulfilled", "rejected"]);
    expect(await readWorkspace()).not.toBeNull();
  });
  it("allows explicitly confirmed replacement and then protects subsequent saves", async () => {
    await writeWorkspace(snapshot());
    const engine = createDemoEngine();
    engine.setWorkbookMetadata("Launch budget", { restored: true });
    const writer = createWorkspaceWriter(null, true);
    await writer.save(engine, view, compact);
    expect(
      (await loadWorkspace())!.engine.getWorkbookMetadata("Launch budget"),
    ).toEqual({ restored: true });
    await writeWorkspace(snapshot());
    await expect(writer.save(engine, view, warm)).rejects.toThrow(
      "Another tab changed",
    );
  });
  it("preserves incomplete chunked data on read failure", async () => {
    await writeWorkspace(snapshot());
    const db = await openDB("spreadsheet-arena", 1);
    const current = await db.get("workspace", "current");
    await db.delete("workspace", `engine:${current.engineId}:0`);
    await expect(readWorkspace()).rejects.toThrow(
      "original data has been kept",
    );
    expect(await db.get("workspace", "current")).toEqual(current);
    db.close();
  });
  it("reports unavailable storage instead of silently falling back", async () => {
    const writer = createWorkspaceWriter();
    const engine = createDemoEngine();
    vi.stubGlobal("indexedDB", undefined);
    try {
      await expect(writer.save(engine, view, warm)).rejects.toThrow(
        "IndexedDB is unavailable",
      );
    } finally {
      vi.unstubAllGlobals();
    }
    await writer.save(engine, view, compact);
    expect(await readWorkspace()).not.toBeNull();
  });
});
