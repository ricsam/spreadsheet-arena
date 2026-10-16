import { openDB } from "idb";
import type { FormulaEngine } from "@ricsam/formula-engine";
import {
  restoreSnapshot,
  validateWorkspaceView,
  type LoadedWorkspace,
  type WorkspaceSnapshot,
  type WorkspaceView,
} from "./model";

const databaseName = "spreadsheet-arena";
const storeName = "workspace";
const key = "current";
// Stay well below Chromium's per-value structured-clone limit, including UTF-16.
export const ENGINE_CHUNK_CHARS = 1024 * 1024;
type StoredWorkspace = WorkspaceView & {
  format: "spreadsheet-arena-storage";
  version: 1;
  engineId: string;
  chunks: number;
};
function isStored(value: unknown): value is StoredWorkspace {
  if (!value || typeof value !== "object") return false;
  const v = value as StoredWorkspace;
  return (
    v.format === "spreadsheet-arena-storage" &&
    v.version === 1 &&
    typeof v.engineId === "string" &&
    v.engineId.length > 0 &&
    Number.isSafeInteger(v.chunks) &&
    v.chunks > 0 &&
    v.chunks <= 4096
  );
}
// A legacy token is used only until the first atomic migration; never reserialize it.
export type StorageToken = string | { legacyEngine: string } | null;
const tokenFor = (stored: unknown): StorageToken =>
  isStored(stored)
    ? stored.engineId
    : stored &&
        typeof stored === "object" &&
        "engine" in stored &&
        typeof stored.engine === "string"
      ? { legacyEngine: stored.engine }
      : null;
function matchesToken(stored: unknown, expected: StorageToken): boolean {
  if (expected === null) return stored === undefined;
  if (typeof expected === "string")
    return isStored(stored) && stored.engineId === expected;
  return (
    !isStored(stored) &&
    !!stored &&
    typeof stored === "object" &&
    "engine" in stored &&
    stored.engine === expected.legacyEngine
  );
}
const chunkKey = (id: string, index: number) => `engine:${id}:${index}`;
async function database() {
  if (typeof indexedDB === "undefined")
    throw new Error(
      "IndexedDB is unavailable. Your workspace cannot be saved in this browser.",
    );
  return openDB(databaseName, 1, {
    upgrade(db) {
      if (!db.objectStoreNames.contains(storeName))
        db.createObjectStore(storeName);
    },
  });
}
let pending: Promise<void> = Promise.resolve();
function enqueue(write: () => Promise<void>): Promise<void> {
  const operation = pending.then(write);
  pending = operation.catch(() => {});
  return operation;
}
async function readStored(): Promise<{
  value: unknown;
  storageToken: StorageToken;
}> {
  await pending;
  const db = await database();
  try {
    const tx = db.transaction(storeName, "readonly");
    const stored: unknown = await tx.store.get(key);
    if (!isStored(stored))
      return { value: stored, storageToken: tokenFor(stored) };
    const pieces = await Promise.all(
      Array.from({ length: stored.chunks }, (_, i) =>
        tx.store.get(chunkKey(stored.engineId, i)),
      ),
    );
    await tx.done;
    if (pieces.some((p) => typeof p !== "string"))
      throw new Error("A saved workbook chunk is missing or unreadable.");
    return {
      storageToken: stored.engineId,
      value: {
        format: "spreadsheet-arena",
        version: 1,
        engine: pieces.join(""),
        layout: stored.layout,
        activePaneId: stored.activePaneId,
        ...(stored.view ? { view: stored.view } : {}),
      },
    };
  } finally {
    db.close();
  }
}
export async function loadWorkspace(): Promise<
  (LoadedWorkspace & { storageToken: StorageToken }) | null
> {
  try {
    const { value, storageToken } = await readStored();
    return value === undefined
      ? null
      : { ...restoreSnapshot(value), storageToken };
  } catch (error) {
    throw new Error(
      `Saved workspace could not be read; the original data has been kept. ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}
export async function readWorkspace(): Promise<WorkspaceSnapshot | null> {
  return (await loadWorkspace())?.snapshot ?? null;
}

/** Atomic content+view commit, or small view-only commit. Never assembles old engine chunks. */
async function persist(
  view: WorkspaceView,
  serialized?: string,
  expected?: StorageToken,
): Promise<string> {
  const db = await database();
  try {
    const tx = db.transaction(storeName, "readwrite");
    try {
      const previous: unknown = await tx.store.get(key);
      if (expected !== undefined && !matchesToken(previous, expected))
        throw new Error(
          "Another tab changed the saved workbook. Download a backup of your changes, then reload before continuing to save.",
        );
      let engineId: string, chunks: number;
      if (serialized !== undefined) {
        engineId = crypto.randomUUID();
        chunks = Math.max(1, Math.ceil(serialized.length / ENGINE_CHUNK_CHARS));
        for (let i = 0; i < chunks; i++)
          await tx.store.put(
            serialized.slice(
              i * ENGINE_CHUNK_CHARS,
              (i + 1) * ENGINE_CHUNK_CHARS,
            ),
            chunkKey(engineId, i),
          );
        if (isStored(previous))
          for (let i = 0; i < previous.chunks; i++)
            await tx.store.delete(chunkKey(previous.engineId, i));
      } else {
        if (!isStored(previous))
          throw new Error("Workbook contents have not been saved yet.");
        ({ engineId, chunks } = previous);
      }
      await tx.store.put(
        {
          format: "spreadsheet-arena-storage",
          version: 1,
          engineId,
          chunks,
          ...view,
        } satisfies StoredWorkspace,
        key,
      );
      await tx.done;
      return engineId;
    } catch (error) {
      try {
        tx.abort();
      } catch {
        /* already completed or aborted */
      }
      await tx.done.catch(() => {});
      throw error;
    }
  } catch (error) {
    throw new Error(
      `Workspace save failed: ${error instanceof Error ? error.message : String(error)}`,
    );
  } finally {
    db.close();
  }
}

/** External snapshot boundary: validate once, before replacing stored data. */
export function writeWorkspace(snapshot: WorkspaceSnapshot): Promise<void> {
  try {
    const loaded = restoreSnapshot(snapshot);
    const view = validateWorkspaceView(loaded.snapshot, loaded.engine);
    return enqueue(async () => {
      await persist(view, loaded.snapshot.engine);
    });
  } catch (error) {
    return Promise.reject(error);
  }
}

/** One writer per mounted workspace; every explicit save captures fresh contents. */
export function createWorkspaceWriter(
  initialToken: StorageToken = null,
  replace = false,
) {
  let committedId: StorageToken | undefined = replace
    ? undefined
    : initialToken;
  return {
    save(
      engine: FormulaEngine,
      state: WorkspaceView,
      options: { includeEvaluationState: boolean },
    ): Promise<void> {
      try {
        const view = validateWorkspaceView(state, engine);
        // Capture synchronously at invocation, before queued writes settle. Both
        // the detached view and engine bytes belong to this explicit save, even
        // when only the save mode or the engine's warm caches have changed.
        const serialized = engine.serializeEngine({
          includeEvaluationState: options.includeEvaluationState,
        });
        return enqueue(async () => {
          // Advance the expected token only after a successful atomic commit.
          committedId = await persist(view, serialized, committedId);
        });
      } catch (error) {
        return Promise.reject(error);
      }
    },
  };
}
