import { test, expect, type Page } from "@playwright/test";
import ExcelJS from "exceljs";
import { FormulaEngine } from "@ricsam/formula-engine";
import { makeSnapshot } from "../src/model";

const launch = (page: Page) =>
  page.getByRole("grid", { name: "Launch budget, Overview", exact: true });
const inputs = (page: Page) =>
  page.getByRole("grid", { name: "Assumptions, Inputs", exact: true });
const gridCell = (grid: ReturnType<Page["getByRole"]>, address: string) =>
  grid.locator(`[data-cell-address="${address}"]`);
async function editViaBar(
  page: Page,
  grid: ReturnType<Page["getByRole"]>,
  address: string,
  value: string,
) {
  await gridCell(grid, address).click();
  await expect(
    page.getByRole("textbox", { name: "Cell address", exact: true }),
  ).toHaveValue(address);
  await page
    .getByRole("textbox", { name: "Cell value or formula", exact: true })
    .fill(value);
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
}

test("connected calculations, native cell editing, undo, and IndexedDB restoration", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto("./");
  await expect(launch(page)).toBeVisible();
  await expect(gridCell(launch(page), "D12")).toHaveText("7,535");
  await editViaBar(page, inputs(page), "B4", "100");
  await expect(gridCell(launch(page), "D12")).toHaveText("12,375");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(gridCell(launch(page), "D12")).toHaveText("7,535");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(gridCell(launch(page), "D12")).toHaveText("12,375");
  await gridCell(inputs(page), "B4").dblclick();
  const editor = page.getByRole("textbox", { name: "Edit B4", exact: true });
  await editor.fill("50");
  await editor.press("Enter");
  await expect(gridCell(launch(page), "D12")).toHaveText("7,975");
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
  await page.reload();
  await expect(gridCell(inputs(page), "B4")).toHaveText("50");
  await expect(gridCell(launch(page), "D12")).toHaveText("7,975");
  expect(errors).toEqual([]);
});

test("create workbook and sheets; split, resize, maximize and close without removing data", async ({
  page,
}) => {
  await page.goto("./");
  await page.getByRole("button", { name: "New workbook", exact: true }).click();
  await page.getByRole("textbox", { name: "Workbook name" }).fill("Planning");
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  const pane = page.getByRole("region", { name: "Planning pane", exact: true });
  await expect(pane).toBeVisible();
  await pane
    .getByRole("button", { name: "Add sheet to Planning", exact: true })
    .click();
  await expect(pane.getByRole("tab", { selected: true })).not.toHaveText(
    "Sheet1",
  );
  await pane.getByRole("button", { name: "Split Planning below" }).click();
  await expect(
    page.getByRole("region", { name: "Planning pane", exact: true }),
  ).toHaveCount(2);
  const separator = page.getByRole("separator", { name: "Resize rows" });
  await separator.focus();
  await separator.press("ArrowDown");
  await expect(separator).toHaveAttribute("aria-valuenow", "55");
  await page.getByRole("button", { name: "Maximize Planning" }).last().click();
  await expect(page.getByRole("grid")).toHaveCount(1);
  await page.getByRole("button", { name: "Restore panes" }).click();
  await expect(page.getByRole("grid")).toHaveCount(4);
  await page
    .getByRole("button", { name: "Close Planning pane", exact: true })
    .last()
    .click();
  await expect(
    page.getByRole("button", { name: "Open Planning", exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
  await page.reload();
  await expect(page.getByRole("grid")).toHaveCount(3);
});

test("Excel and CSV import, corrupt input, workbook export and workspace restoration", async ({
  page,
}) => {
  await page.goto("./");
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Numbers");
  sheet.addRow(["Count", "Total"]);
  sheet.addRow([3, { formula: "A2*2", result: 6 }]);
  await page.getByLabel("Import workbook files").setInputFiles([
    {
      name: "Example.xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: Buffer.from(await workbook.xlsx.writeBuffer()),
    },
  ]);
  const imported = page.getByRole("grid", {
    name: "Example, Numbers",
    exact: true,
  });
  await expect(gridCell(imported, "B2")).toHaveText("6");
  await page.getByLabel("Import workbook files").setInputFiles([
    {
      name: "broken.xlsx",
      mimeType: "application/octet-stream",
      buffer: Buffer.from("not an Excel file"),
    },
  ]);
  await expect(page.getByRole("alert")).toContainText("broken.xlsx");
  await expect(gridCell(imported, "B2")).toHaveText("6");
  const workbookDownload = page.waitForEvent("download");
  await page
    .getByRole("button", { name: "Download Example", exact: true })
    .click();
  const downloaded = await workbookDownload;
  const path = await downloaded.path();
  await page.getByLabel("Import workbook files").setInputFiles({
    name: "copy.json",
    mimeType: "application/json",
    buffer: await (await import("node:fs/promises")).readFile(path!),
  });
  await expect(
    page.getByRole("button", { name: "Open Example (2)", exact: true }),
  ).toBeVisible();
  await page.getByLabel("Import workbook files").setInputFiles([
    {
      name: "People.csv",
      mimeType: "text/csv",
      buffer: Buffer.from("Name,Count\nAda,12\n"),
    },
  ]);
  await expect(
    page.getByRole("button", { name: "Open People", exact: true }),
  ).toBeVisible();
  await page
    .getByText("Workspace", { exact: true })
    .filter({ has: page.locator("svg") })
    .click();
  const backupDownload = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download workspace backup" }).click();
  const backup = await backupDownload;
  const backupBytes = await (
    await import("node:fs/promises")
  ).readFile((await backup.path())!);
  await page.getByLabel("Restore workspace file").setInputFiles({
    name: "backup.json",
    mimeType: "application/json",
    buffer: backupBytes,
  });
  await expect(page.getByRole("dialog")).toContainText("Restore workspace?");
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  await expect(
    page.getByRole("button", { name: "Open People", exact: true }),
  ).toBeVisible();
});

test("WebMCP registration and semantic editing contracts work without browser experimental API", async ({
  page,
}) => {
  await page.addInitScript(() => {
    const tools = new Map();
    Object.defineProperty(navigator, "modelContext", {
      configurable: true,
      value: {
        registerTool: (tool: { name: string }) => tools.set(tool.name, tool),
        unregisterTool: (name: string) => tools.delete(name),
      },
    });
    Object.assign(window, { arenaTestTools: tools });
  });
  await page.goto("./");
  await expect(launch(page)).toBeVisible();
  const result = await page.evaluate(async () => {
    const tools = (
      window as unknown as {
        arenaTestTools: Map<
          string,
          { execute: (args: object) => Promise<unknown> }
        >;
      }
    ).arenaTestTools;
    const list = await tools.get("list_workbooks")!.execute({});
    const set = await tools.get("set_cell")!.execute({
      workbook: "Assumptions",
      sheet: "Inputs",
      cell: "B4",
      value: 100,
    });
    const read = await tools.get("read_cells")!.execute({
      workbook: "Launch budget",
      sheet: "Overview",
      cells: ["D12"],
    });
    return { list, set, read };
  });
  expect(JSON.stringify(result.read)).toContain("12375");
  await expect(gridCell(launch(page), "D12")).toHaveText("12,375");
  await expect(inputs(page)).toHaveAttribute(
    "aria-label",
    "Assumptions, Inputs",
  );
  await gridCell(inputs(page), "B4").click();
  await expect(gridCell(inputs(page), "B4")).toHaveAttribute(
    "aria-selected",
    "true",
  );
});

test("file drop and far-away address navigation are accessible", async ({
  page,
}) => {
  await page.goto("./");
  await expect(launch(page)).toBeVisible();
  const transfer = await page.evaluateHandle(() => {
    const transfer = new DataTransfer();
    transfer.items.add(
      new File(["Name,Total\nDrop,8"], "Dropped.csv", { type: "text/csv" }),
    );
    return transfer;
  });
  await page
    .locator(".app-shell")
    .dispatchEvent("drop", { dataTransfer: transfer });
  await expect(
    page.getByRole("button", { name: "Open Dropped", exact: true }),
  ).toBeVisible();
  await page
    .getByRole("textbox", { name: "Cell address", exact: true })
    .fill("Z100");
  await page.getByRole("button", { name: "Go to cell", exact: true }).click();
  const grid = page.getByRole("grid", { name: /^Dropped,/ });
  await expect(gridCell(grid, "Z100")).toBeVisible();
  await page
    .getByRole("textbox", { name: "Cell value or formula", exact: true })
    .fill("=2+3");
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
  await expect(gridCell(grid, "Z100")).toHaveText("5");
});

test("unreadable IndexedDB state is preserved instead of silently replaced", async ({
  page,
}) => {
  await page.goto("./");
  await expect(launch(page)).toBeVisible();
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
  await page.evaluate(async () => {
    await new Promise<void>((resolve, reject) => {
      const request = indexedDB.open("spreadsheet-arena", 1);
      request.onsuccess = () => {
        const db = request.result;
        const tx = db.transaction("workspace", "readwrite");
        tx.objectStore("workspace").put({ corrupt: "keep me" }, "current");
        tx.oncomplete = () => {
          db.close();
          resolve();
        };
        tx.onerror = () => reject(tx.error);
      };
    });
  });
  await page.reload();
  await expect(page.getByRole("alert")).toContainText(
    "It has not been overwritten",
  );
  await page
    .getByRole("button", { name: "Create a workbook", exact: true })
    .click();
  const saved = await page.evaluate(
    async () =>
      await new Promise((resolve) => {
        const request = indexedDB.open("spreadsheet-arena", 1);
        request.onsuccess = () => {
          const db = request.result;
          const get = db
            .transaction("workspace")
            .objectStore("workspace")
            .get("current");
          get.onsuccess = () => {
            resolve(get.result);
            db.close();
          };
        };
      }),
  );
  expect(saved).toEqual({ corrupt: "keep me" });
});

test("navigation persists only view state, while edits replace workbook content", async ({
  page,
}) => {
  await page.goto("./");
  const saved = page
    .getByRole("status")
    .filter({ hasText: "Saved on this device" });
  await expect(saved).toBeVisible();
  const stored = () =>
    page.evaluate(
      async () =>
        new Promise<Record<string, any>>((resolve, reject) => {
          const request = indexedDB.open("spreadsheet-arena", 1);
          request.onsuccess = () => {
            const db = request.result;
            const get = db
              .transaction("workspace")
              .objectStore("workspace")
              .get("current");
            get.onsuccess = () => {
              resolve(get.result);
              db.close();
            };
            get.onerror = () => reject(get.error);
          };
        }),
    );
  const initial = await stored();
  expect(initial.format).toBe("spreadsheet-arena-storage");
  await gridCell(launch(page), "A4").click();
  await page.keyboard.press("ArrowDown");
  await expect(
    page.getByRole("textbox", { name: "Cell address", exact: true }),
  ).toHaveValue("A5");
  await expect(saved).toBeVisible();
  const moved = await stored();
  expect(moved.engineId).toBe(initial.engineId);
  await editViaBar(page, inputs(page), "B4", "75");
  await expect(saved).toBeVisible();
  expect((await stored()).engineId).not.toBe(initial.engineId);
  await page.reload();
  await expect(gridCell(inputs(page), "B4")).toHaveText("75");
});

test("restores backups above the old 25 MiB cap and persists content in bounded chunks", async ({
  page,
}) => {
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: "Large backup",
    data: {
      workbookMetadata: { note: "x".repeat(26 * 1024 * 1024) },
      sheets: [{ name: "Sheet1", content: { A1: "Preserved", B1: "=21*2" } }],
    },
  });
  const snapshot = makeSnapshot(
    engine,
    {
      kind: "pane",
      pane: {
        id: "large",
        workbookName: "Large backup",
        sheetName: "Sheet1",
        cell: "A1",
      },
    },
    "large",
  );
  await page.goto("./");
  await expect(launch(page)).toBeVisible();
  await page.getByLabel("Restore workspace file").setInputFiles({
    name: "large.workspace.json",
    mimeType: "application/json",
    buffer: Buffer.from(JSON.stringify(snapshot)),
  });
  await expect(page.getByRole("dialog")).toContainText("Restore workspace?");
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  const grid = page.getByRole("grid", {
    name: "Large backup, Sheet1",
    exact: true,
  });
  await expect(gridCell(grid, "B1")).toHaveText("42");
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
  await page.reload();
  await expect(gridCell(grid, "A1")).toHaveText("Preserved");
  await expect(gridCell(grid, "B1")).toHaveText("42");
});

test("Excel default and inherited fonts match explicit cells and survive reload", async ({
  page,
}) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Fonts");
  sheet.getCell("B13").value = "same";
  sheet.getCell("B13").font = {
    name: "Calibri",
    family: 2,
    size: 11,
    color: { theme: 1 },
    scheme: "minor",
  };
  sheet.getCell("B14").value = "same";
  sheet.getRow(16).font = {
    name: "Arial",
    size: 12,
    bold: true,
    color: { argb: "FF000000" },
  };
  sheet.getCell("B16").value = "inherited";
  await page.goto("./");
  await page
    .getByLabel("Import workbook files")
    .setInputFiles({
      name: "Fonts.xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: Buffer.from(await workbook.xlsx.writeBuffer()),
    });
  const grid = page.getByRole("grid", { name: "Fonts, Fonts", exact: true });
  await expect(grid).toBeVisible();
  await page
    .getByRole("textbox", { name: "Cell address", exact: true })
    .fill("B14");
  await page.getByRole("button", { name: "Go to cell", exact: true }).click();
  for (const cell of ["B13", "B14", "B15"]) {
    await expect(gridCell(grid, cell)).toHaveCSS("font-size", "15px");
    await expect(gridCell(grid, cell)).toHaveCSS("color", "rgb(0, 0, 0)");
    await expect(gridCell(grid, cell)).toHaveCSS("font-family", "Calibri");
  }
  await expect(gridCell(grid, "B16")).toHaveCSS("font-size", "16px");
  await expect(gridCell(grid, "B16")).toHaveCSS("font-weight", "700");
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
  await page.reload();
  await expect(gridCell(grid, "B14")).toHaveCSS("font-size", "15px");
  await expect(gridCell(grid, "B16")).toHaveCSS("font-weight", "700");
});

test("mobile toolbar and workbook creation remain usable", async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("./");
  await expect(
    page.getByRole("button", { name: "Show sidebar" }),
  ).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(
    390,
  );
  await page.getByRole("button", { name: "Help and automation" }).click();
  await expect(
    page.getByRole("region", { name: "Help and automation" }),
  ).toBeVisible();
  await page.getByRole("button", { name: "Close help" }).click();
  await page.getByRole("button", { name: "Show sidebar" }).click();
  await page.getByRole("button", { name: "New workbook", exact: true }).click();
  await page.getByRole("textbox", { name: "Workbook name" }).fill("Mobile");
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  await expect(
    page.getByRole("button", { name: "Open Mobile", exact: true }),
  ).toBeVisible();
});
