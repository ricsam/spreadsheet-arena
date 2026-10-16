import { test, expect, type Page } from "@playwright/test";

const input = (page: Page) =>
  page
    .getByRole("grid", { name: "Assumptions, Inputs", exact: true })
    .locator('[data-cell-address="B4"]');
const status = (page: Page, text: string) =>
  page.getByRole("status").filter({ hasText: text });
async function edit(page: Page, value: string) {
  await input(page).click();
  await page
    .getByRole("textbox", { name: "Cell value or formula", exact: true })
    .fill(value);
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
}
async function save(page: Page) {
  await page.getByRole("button", { name: "Save", exact: true }).click();
  await expect(status(page, "Saved on this device")).toBeVisible();
}
async function stored(page: Page) {
  return page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open("spreadsheet-arena", 1);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
    try {
      if (!db.objectStoreNames.contains("workspace")) return null;
      const tx = db.transaction("workspace");
      const get = (key: string) =>
        new Promise<any>((resolve, reject) => {
          const r = tx.objectStore("workspace").get(key);
          r.onsuccess = () => resolve(r.result);
          r.onerror = () => reject(r.error);
        });
      const current = await get("current");
      if (!current) return null;
      const chunks = await Promise.all(
        Array.from({ length: current.chunks }, (_, i) =>
          get(`engine:${current.engineId}:${i}`),
        ),
      );
      return {
        id: current.engineId,
        snapshot: JSON.parse(chunks.join("")),
        layout: current.layout,
      };
    } finally {
      db.close();
    }
  });
}

test("saves only explicitly, offers both cache modes and protects unsaved navigation", async ({
  page,
}) => {
  await page.goto("./");
  await expect(input(page)).toHaveText("45");
  await expect(status(page, "Unsaved changes")).toBeVisible();
  await expect(
    page.getByRole("combobox", { name: "Save contents" }),
  ).toHaveValue("compact");
  expect(await stored(page)).toBeNull();
  await edit(page, "77");
  await page.waitForTimeout(800); // Regression: the old 350ms autosave must not run.
  expect(await stored(page)).toBeNull();
  await page.keyboard.press("Control+s");
  await expect(status(page, "Saved on this device")).toBeVisible();
  const compact = (await stored(page))!;
  expect(compact.snapshot.managers.dependency).toBeUndefined();
  expect(compact.snapshot.managers.cache).toBeUndefined();
  await page.reload();
  await expect(input(page)).toHaveText("77");
  await page
    .getByRole("combobox", { name: "Save contents" })
    .selectOption("warm");
  await save(page);
  const warm = (await stored(page))!;
  expect(warm.id).not.toBe(compact.id);
  expect(warm.snapshot.managers.dependency).toBeDefined();
  expect(warm.snapshot.managers.cache).toBeDefined();
  await page
    .getByRole("combobox", { name: "Save contents" })
    .selectOption("compact");
  await save(page); // Same revision, mode change still writes the selected representation.
  expect((await stored(page))!.snapshot.managers.dependency).toBeUndefined();
  await edit(page, "88");
  const before = (await stored(page))!.id;
  let warned = false;
  page.once("dialog", (dialog) => {
    warned = dialog.type() === "beforeunload";
    return dialog.dismiss();
  });
  // A cancelled navigation has no load event; bound the reload observer, not
  // the rest of this test, while still verifying the real beforeunload prompt.
  await page.reload({ timeout: 1500 }).catch(() => {});
  expect(warned).toBe(true);
  await expect(input(page)).toHaveText("88");
  expect((await stored(page))!.id).toBe(before);
  await save(page);
  await page.reload();
  await expect(input(page)).toHaveText("88");
});

test("failed saves unlock editing, retain old data and retry successfully", async ({
  page,
}) => {
  await page.goto("./");
  await expect(input(page)).toBeVisible();
  await save(page);
  const old = (await stored(page))!.id;
  await edit(page, "91");
  await page.evaluate(() => {
    const original = IDBObjectStore.prototype.put;
    (window as any).__restorePut = () => {
      IDBObjectStore.prototype.put = original;
    };
    IDBObjectStore.prototype.put = function (value, key) {
      if (key === "current")
        throw new DOMException("test quota", "QuotaExceededError");
      return original.call(this, value, key);
    };
  });
  await page.getByRole("button", { name: "Save", exact: true }).click();
  await expect(page.getByRole("alert")).toContainText("test quota");
  await expect(status(page, "Save failed — unsaved changes")).toBeVisible();
  await expect(
    page.getByRole("dialog", { name: "Saving workspace…" }),
  ).toHaveCount(0);
  expect((await stored(page))!.id).toBe(old);
  await edit(page, "92");
  await page.evaluate(() => (window as any).__restorePut());
  await save(page);
  await page.reload();
  await expect(input(page)).toHaveText("92");
});

test("save locks controls, keyboard and previously registered automation until commit", async ({
  page,
}) => {
  await page.addInitScript(() => {
    const tools = new Map();
    (window as any).__tools = tools;
    Object.defineProperty(navigator, "modelContext", {
      configurable: true,
      value: {
        registerTool: (tool: any) => tools.set(tool.name, tool),
        unregisterTool: (name: string) => tools.delete(name),
      },
    });
    const original = IDBObjectStore.prototype.put;
    IDBObjectStore.prototype.put = function (value, key) {
      const request = original.call(this, value, key);
      if (key === "current" && (window as any).__holdSave) {
        const store = this;
        const keepAlive = () => {
          if ((window as any).__holdSave)
            store.get("current").onsuccess = keepAlive;
        };
        keepAlive();
      }
      return request;
    };
  });
  await page.goto("./");
  await expect(input(page)).toBeVisible();
  await page.evaluate(() => {
    (window as any).__holdSave = true;
  });
  await page.getByRole("button", { name: "Save", exact: true }).click();
  await expect(
    page.getByRole("dialog", { name: "Saving workspace…" }),
  ).toBeVisible();
  await expect(page.locator(".app-body")).toHaveAttribute("inert", "");
  await expect(page.getByLabel("Import workbook files")).toBeDisabled();
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+z");
  await page.keyboard.press("9");
  const result = await page.evaluate(async () =>
    (window as any).__tools.get("set_cell").execute({
      workbook: "Assumptions",
      sheet: "Inputs",
      cell: "B4",
      value: 999,
    }),
  );
  expect(result.isError).toBe(true);
  expect(result.content[0].text).toContain("Saving in progress");
  await expect(
    page.getByRole("dialog", { name: "Saving workspace…" }),
  ).toBeVisible();
  await page.evaluate(() => {
    (window as any).__holdSave = false;
  });
  await expect(status(page, "Saved on this device")).toBeVisible();
  await expect(input(page)).toHaveText("45");
  await edit(page, "46");
  await expect(input(page)).toHaveText("46");
});

test("a clipboard read started before saving cannot mutate the locked workspace", async ({
  page,
}) => {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, "clipboard", {
      configurable: true,
      value: {
        readText: () =>
          new Promise<string>((resolve) => {
            (window as any).__resolveClipboard = resolve;
          }),
      },
    });
    const original = IDBObjectStore.prototype.put;
    IDBObjectStore.prototype.put = function (value, key) {
      const request = original.call(this, value, key);
      if (key === "current" && (window as any).__holdSave) {
        const store = this;
        const keepAlive = () => {
          if ((window as any).__holdSave)
            store.get("current").onsuccess = keepAlive;
        };
        keepAlive();
      }
      return request;
    };
  });
  await page.goto("./");
  await input(page).click();
  await page.getByRole("button", { name: "Paste", exact: true }).click();
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).__resolveClipboard))
    .toBe("function");
  await page.evaluate(() => {
    (window as any).__holdSave = true;
  });
  await page.getByRole("button", { name: "Save", exact: true }).click();
  await expect(
    page.getByRole("dialog", { name: "Saving workspace…" }),
  ).toBeVisible();
  await page.evaluate(async () => {
    (window as any).__resolveClipboard("999");
    await new Promise((resolve) => setTimeout(resolve, 0));
    (window as any).__holdSave = false;
  });
  await expect(status(page, "Saved on this device")).toBeVisible();
  await expect(input(page)).toHaveText("45");
  await page.reload();
  await expect(input(page)).toHaveText("45");
});

test("save shortcut refuses uncommitted formula-bar and cell-editor drafts", async ({
  page,
}) => {
  await page.goto("./");
  await expect(input(page)).toBeVisible();
  await save(page);
  await input(page).click();
  const bar = page.getByRole("textbox", {
    name: "Cell value or formula",
    exact: true,
  });
  await bar.fill("99");
  await bar.press("Control+s");
  await expect(page.getByRole("alert")).toContainText("Apply or cancel");
  await expect(bar).toHaveValue("99");
  await bar.press("Escape");
  await input(page).dblclick();
  const editor = page.getByRole("textbox", { name: "Edit B4", exact: true });
  await editor.fill("88");
  await editor.press("Control+s");
  await expect(editor).toHaveValue("88");
  await expect(
    page.getByRole("dialog", { name: "Saving workspace…" }),
  ).toHaveCount(0);
  await editor.press("Enter");
  await save(page);
  await page.reload();
  await expect(input(page)).toHaveText("88");
});

test("inspector drafts block save shortcuts and warn before leaving", async ({
  page,
}) => {
  await page.goto("./");
  await input(page).click();
  await save(page);
  const fontSize = page.getByRole("spinbutton", { name: "Font size" });
  await fontSize.fill("22");
  await fontSize.press("Control+s");
  await expect(page.getByRole("alert")).toContainText("Apply or cancel");
  await expect(fontSize).toHaveValue("22");
  expect(
    await page.evaluate(() => {
      const event = new Event("beforeunload", { cancelable: true });
      window.dispatchEvent(event);
      return event.defaultPrevented;
    }),
  ).toBe(true);
  await fontSize.press("Enter");
  await save(page);
  await page.reload();
  await expect(input(page)).toHaveCSS("font-size", "22px");
});

test("an import keeps pane changes made while parsing and cannot overlap save", async ({
  page,
}) => {
  await page.addInitScript(() => {
    const original = File.prototype.text;
    File.prototype.text = async function () {
      if (this.name === "slow.json")
        await new Promise<void>((resolve) => {
          (window as any).__finishImport = resolve;
        });
      return original.call(this);
    };
  });
  await page.goto("./");
  await expect(input(page)).toBeVisible();
  await page.getByLabel("Import workbook files").setInputFiles({
    name: "slow.json",
    mimeType: "application/json",
    buffer: Buffer.from(
      JSON.stringify({
        name: "Imported",
        sheets: [{ name: "Data", content: { A1: 1 } }],
      }),
    ),
  });
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).__finishImport))
    .toBe("function");
  await expect(
    page.getByRole("button", { name: "Save", exact: true }),
  ).toBeDisabled();
  await page
    .getByRole("button", { name: "Split Assumptions below", exact: true })
    .click();
  await expect(page.getByRole("grid")).toHaveCount(3);
  await page.evaluate(() => (window as any).__finishImport());
  await expect(page.getByRole("grid")).toHaveCount(4);
  await expect(
    page.getByRole("region", { name: "Assumptions pane", exact: true }),
  ).toHaveCount(2);
  await save(page);
});

test("manual save detects another tab's save without overwriting it", async ({
  page,
  context,
}) => {
  await page.goto("./");
  await expect(input(page)).toBeVisible();
  await save(page);
  const other = await context.newPage();
  await other.goto("./");
  await expect(input(other)).toHaveText("45");
  await edit(page, "61");
  await save(page);
  await edit(other, "62");
  await other.getByRole("button", { name: "Save", exact: true }).click();
  await expect(other.getByRole("alert")).toContainText("Another tab changed");
  await expect(input(other)).toHaveText("62");
  await page.reload();
  await expect(input(page)).toHaveText("61");
});

test("save is available on mobile without opening the sidebar", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("./");
  await expect(
    page.getByRole("button", { name: "Save", exact: true }),
  ).toBeInViewport();
  await expect(
    page.getByRole("combobox", { name: "Save contents" }),
  ).toBeInViewport();
  await save(page);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(
    390,
  );
});
