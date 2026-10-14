import "fake-indexeddb/auto";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { deleteDB, openDB } from "idb";
import { createDemoEngine, makeSnapshot } from "./model";
import { readWorkspace, writeWorkspace } from "./persistence";

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
