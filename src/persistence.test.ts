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

describe("IndexedDB persistence", () => {
  it("returns null only when no saved workspace exists", async () => {
    expect(await readWorkspace()).toBeNull();
  });
  it("saves and restores the native engine snapshot", async () => {
    const value = snapshot();
    await writeWorkspace(value);
    expect(await readWorkspace()).toEqual(value);
  });
  it("serializes concurrent writes in invocation order and detaches mutable layouts", async () => {
    const a = snapshot(),
      b = snapshot();
    b.layout = {
      kind: "pane",
      pane: {
        id: "last",
        workbookName: "Launch budget",
        sheetName: "Overview",
        cell: "D12",
        zoom: 1.25,
      },
    };
    b.activePaneId = "last";
    const expected = structuredClone(b);
    const one = writeWorkspace(a),
      two = writeWorkspace(b);
    b.layout.pane.cell = "A1";
    await Promise.all([one, two]);
    expect(await readWorkspace()).toEqual(expected);
  });
  it("does not erase corrupted stored data on read", async () => {
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
  it("loads deployed single-record snapshots without a destructive migration", async () => {
    await readWorkspace();
    const db = await openDB("spreadsheet-arena", 1);
    const value = snapshot();
    await db.put("workspace", value, "current");
    const reset = vi.spyOn(FormulaEngine.prototype, "resetToSerializedEngine");
    const loaded = await loadWorkspace();
    expect(reset).toHaveBeenCalledTimes(1);
    reset.mockRestore();
    expect(loaded?.snapshot).toEqual(value);
    expect(loaded?.engine.getWorkbooks().size).toBe(2);
    expect(await db.get("workspace", "current")).toEqual(value);
    db.close();
  });
  it("atomically migrates a loaded legacy record without rebuilding or serializing it", async () => {
    await readWorkspace();
    const db = await openDB("spreadsheet-arena", 1);
    const initial = snapshot();
    await db.put("workspace", initial, "current");
    const loaded = (await loadWorkspace())!;
    const writer = createWorkspaceWriter(loaded.storageToken, false, {
      engine: loaded.engine,
      revision: 0,
      serialized: loaded.snapshot.engine,
    });
    const serialize = vi.spyOn(loaded.engine, "serializeEngine");
    await writer.save(loaded.engine, 0, loaded.snapshot);
    expect(serialize).not.toHaveBeenCalled();
    expect((await db.get("workspace", "current")).format).toBe(
      "spreadsheet-arena-storage",
    );
    expect((await readWorkspace())?.engine).toBe(initial.engine);
    serialize.mockRestore();
    db.close();
  });
  it("reuses loaded warm data without serializing on first save or selection", async () => {
    await writeWorkspace(snapshot());
    const loaded = (await loadWorkspace())!;
    const writer = createWorkspaceWriter(loaded.storageToken, false, {
      engine: loaded.engine,
      revision: 0,
      serialized: loaded.snapshot.engine,
    });
    const serialize = vi.spyOn(loaded.engine, "serializeEngine");
    await writer.save(loaded.engine, 0, loaded.snapshot);
    expect(serialize).not.toHaveBeenCalled();
    serialize.mockRestore();
  });
  it("serializes only content revisions and writes only small metadata on navigation", async () => {
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
    const reset = vi.spyOn(FormulaEngine.prototype, "resetToSerializedEngine");
    await writer.save(engine, 0, state);
    const db = await openDB("spreadsheet-arena", 1);
    const first = await db.get("workspace", "current");
    const keys = await db.getAllKeys("workspace");
    const moves = Array.from({ length: 12 }, (_, i) =>
      writer.save(engine, 0, {
        ...state,
        layout: {
          kind: "pane",
          pane: {
            id: "p",
            workbookName: "Launch budget",
            sheetName: "Overview",
            cell: `B${i + 1}`,
          },
        },
      }),
    );
    await Promise.all(moves);
    expect(serialize).toHaveBeenCalledTimes(1);
    expect(serialize).toHaveBeenCalledWith();
    expect(reset).not.toHaveBeenCalled();
    const moved = await db.get("workspace", "current");
    expect(moved.engineId).toBe(first.engineId);
    expect(moved.layout.pane.cell).toBe("B12");
    expect(await db.getAllKeys("workspace")).toEqual(keys);
    engine.setCellContent(
      {
        workbookName: "Assumptions",
        sheetName: "Inputs",
        rowIndex: 3,
        colIndex: 1,
      },
      60,
    );
    await writer.save(engine, 1, state);
    expect(serialize).toHaveBeenCalledTimes(2);
    const edited = await db.get("workspace", "current");
    expect(edited.engineId).not.toBe(first.engineId);
    expect(
      (await db.getAllKeys("workspace")).some((k) =>
        String(k).includes(first.engineId),
      ),
    ).toBe(false);
    reset.mockRestore();
    serialize.mockRestore();
    db.close();
    expect(
      (await loadWorkspace())?.engine.getCellValue({
        workbookName: "Assumptions",
        sheetName: "Inputs",
        rowIndex: 3,
        colIndex: 1,
      }),
    ).toBe(60);
  });
  it("chunks large content and preserves old data if a transaction fails", async () => {
    const engine = createDemoEngine();
    engine.setWorkbookMetadata("Launch budget", {
      payload: "x".repeat(ENGINE_CHUNK_CHARS * 2),
    });
    const writer = createWorkspaceWriter();
    const view = { layout: null, activePaneId: null };
    await writer.save(engine, 0, view);
    const db = await openDB("spreadsheet-arena", 1);
    const before = await db.get("workspace", "current");
    expect(before.chunks).toBeGreaterThan(2);
    for (const key of await db.getAllKeys("workspace")) {
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
    await expect(writer.save(engine, 1, view)).rejects.toThrow("full");
    fail.mockRestore();
    expect(await db.get("workspace", "current")).toEqual(before);
    const loaded = await loadWorkspace();
    expect(loaded?.engine.getWorkbookMetadata("Launch budget")).toEqual({
      payload: "x".repeat(ENGINE_CHUNK_CHARS * 2),
    });
    // Failed content must be retried even when the revision is unchanged.
    await writer.save(engine, 1, view);
    expect(
      (await loadWorkspace())?.engine.getWorkbookMetadata("Launch budget"),
    ).toEqual({ payload: "changed" });
    db.close();
  });
  it("does not attach stale tab view state to another tab's workbook contents", async () => {
    const engine = createDemoEngine(),
      writer = createWorkspaceWriter();
    const view = { layout: null, activePaneId: null };
    await writer.save(engine, 0, view);
    const other = createDemoEngine();
    other.setWorkbookMetadata("Launch budget", { newerTab: true });
    await writeWorkspace(makeSnapshot(other, null, null));
    await expect(writer.save(engine, 0, view)).rejects.toThrow(
      "Another tab changed",
    );
    engine.setWorkbookMetadata("Launch budget", { staleEdit: true });
    await expect(writer.save(engine, 1, view)).rejects.toThrow(
      "Another tab changed",
    );
    expect(
      (await loadWorkspace())?.engine.getWorkbookMetadata("Launch budget"),
    ).toEqual({ newerTab: true });
  });
  it("guards first autosave after load, including legacy migration, against intervening writes", async () => {
    for (const legacy of [false, true]) {
      const initial = snapshot();
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
      await expect(
        writer.save(loaded.engine, 0, loaded.snapshot),
      ).rejects.toThrow("Another tab changed");
      expect(
        (await loadWorkspace())?.engine.getWorkbookMetadata("Launch budget"),
      ).toEqual({ newer: true });
    }
  });
  it("only allows one fresh writer to claim an initially empty workspace", async () => {
    const first = createWorkspaceWriter(),
      second = createWorkspaceWriter();
    const view = { layout: null, activePaneId: null };
    const results = await Promise.allSettled([
      first.save(createDemoEngine(), 0, view),
      second.save(createDemoEngine(), 0, view),
    ]);
    expect(results.map((r) => r.status)).toEqual(["fulfilled", "rejected"]);
    expect(await readWorkspace()).not.toBeNull();
  });
  it("allows explicitly confirmed replacement and then protects subsequent saves", async () => {
    await writeWorkspace(snapshot());
    const engine = createDemoEngine();
    engine.setWorkbookMetadata("Launch budget", { restored: true });
    const writer = createWorkspaceWriter(null, true);
    await writer.save(engine, 0, { layout: null, activePaneId: null });
    expect(
      (await loadWorkspace())?.engine.getWorkbookMetadata("Launch budget"),
    ).toEqual({ restored: true });
    await writeWorkspace(snapshot());
    await expect(
      writer.save(engine, 1, { layout: null, activePaneId: null }),
    ).rejects.toThrow("Another tab changed");
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
    vi.stubGlobal("indexedDB", undefined);
    try {
      await expect(writeWorkspace(snapshot())).rejects.toThrow(
        "IndexedDB is unavailable",
      );
    } finally {
      vi.unstubAllGlobals();
    }
    await writeWorkspace(snapshot());
    expect(await readWorkspace()).not.toBeNull();
  });
});
