import { openDB } from "idb";
import { validateSnapshot, type WorkspaceSnapshot } from "./model";

const databaseName = "spreadsheet-arena";
const storeName = "workspace";
const key = "current";
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
export async function readWorkspace(): Promise<WorkspaceSnapshot | null> {
  await pending;
  const db = await database();
  try {
    const value: unknown = await db.get(storeName, key);
    if (value === undefined) return null;
    try {
      return validateSnapshot(value);
    } catch (error) {
      throw new Error(
        `Saved workspace could not be read; the original data has been kept. ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  } finally {
    db.close();
  }
}
export function writeWorkspace(snapshot: WorkspaceSnapshot): Promise<void> {
  // Capture a validated, detached value now, not when an earlier write completes.
  let value: WorkspaceSnapshot;
  try {
    value = validateSnapshot(snapshot);
  } catch (error) {
    return Promise.reject(error);
  }
  const operation = pending.then(async () => {
    const db = await database();
    try {
      await db.put(storeName, value, key);
    } catch (error) {
      throw new Error(
        `Workspace save failed: ${error instanceof Error ? error.message : String(error)}`,
      );
    } finally {
      db.close();
    }
  });
  // Failure is exposed to this caller, but must not poison later saves.
  pending = operation.catch(() => {});
  return operation;
}
