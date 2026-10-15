import { test, expect, type Page } from "@playwright/test";
import { FormulaEngine } from "@ricsam/formula-engine";
import { makeSnapshot } from "../src/model";

const gridCell = (page: Page, grid: string, address: string) =>
  page
    .getByRole("grid", { name: grid, exact: true })
    .locator(`[data-cell-address="${address}"]`);
const formulaBar = (page: Page) =>
  page.getByRole("textbox", { name: "Cell value or formula", exact: true });
const saved = (page: Page) =>
  page.getByRole("status").filter({ hasText: "Saved on this device" });
async function enableDebug(page: Page) {
  await expect(page.getByRole("main")).toBeVisible();
  const showSidebar = page.getByRole("button", {
    name: "Show sidebar",
    exact: true,
  });
  if (await showSidebar.isVisible()) await showSidebar.click();
  await page.getByRole("button", { name: "Workspace", exact: true }).click();
  const toggle = page.getByRole("menuitemcheckbox", {
    name: "Debug mode",
    exact: true,
  });
  await expect(toggle).toHaveAttribute("aria-checked", "false");
  await toggle.click();
  await expect(page.getByRole("region", { name: "Debug tools" })).toBeVisible();
  if ((page.viewportSize()?.width ?? 1440) <= 900)
    await page
      .getByRole("button", { name: "Hide sidebar", exact: true })
      .click();
}
async function clearCache(page: Page) {
  await page
    .getByRole("button", { name: "Clear formula cache", exact: true })
    .click();
  await expect(
    page.getByRole("status").filter({ hasText: "Formula cache cleared." }),
  ).toBeVisible();
}

test("debug cache clearing preserves panes, selection, drafts, content and undo/redo", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("./");
  const total = gridCell(page, "Launch budget, Overview", "D12");
  const input = gridCell(page, "Assumptions, Inputs", "B4");
  await expect(total).toHaveText("7,535");
  await expect(
    page.getByRole("button", { name: "Clear formula cache" }),
  ).toHaveCount(0);
  await input.click();
  await formulaBar(page).fill("100");
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
  await expect(total).toHaveText("12,375");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(total).toHaveText("7,535");
  await expect(
    page.getByRole("button", { name: "Redo", exact: true }),
  ).toBeEnabled();
  const paneIds = await page
    .locator("[data-pane-id]")
    .evaluateAll((panes) =>
      panes.map((pane) => pane.getAttribute("data-pane-id")),
    );
  const inputElement = await input.elementHandle();
  await enableDebug(page);
  await formulaBar(page).fill("Uncommitted draft");
  await clearCache(page);
  await clearCache(page); // Repeated clearing is safe, not a new workbook edit.
  await expect(total).toHaveText("7,535");
  await expect(input).toHaveText("45");
  await expect(formulaBar(page)).toHaveValue("Uncommitted draft");
  await expect(
    page.getByRole("textbox", { name: "Cell address", exact: true }),
  ).toHaveValue("B4");
  await expect(input).toHaveAttribute("aria-selected", "true");
  expect(await inputElement!.evaluate((node) => node.isConnected)).toBe(true);
  expect(
    await page
      .locator("[data-pane-id]")
      .evaluateAll((panes) =>
        panes.map((pane) => pane.getAttribute("data-pane-id")),
      ),
  ).toEqual(paneIds);
  // Cancel only the uncommitted draft; cache clearing must not consume redo.
  await formulaBar(page).press("Escape");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(total).toHaveText("12,375");
  await clearCache(page);
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(total).toHaveText("7,535");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(saved(page)).toBeVisible();
  await page.reload();
  await expect(total).toHaveText("12,375");
  await expect(input).toHaveText("100");
  await expect(page.getByRole("region", { name: "Debug tools" })).toHaveCount(
    0,
  );
  expect(errors).toEqual([]);
});

test("clear cache replaces stale restored calculations and refreshes mounted cells", async ({
  page,
}) => {
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: "Cache test",
    data: { sheets: [{ name: "Sheet1", content: { A1: 21, B1: "=A1*2" } }] },
  });
  expect(
    engine.getCellValue({
      workbookName: "Cache test",
      sheetName: "Sheet1",
      rowIndex: 0,
      colIndex: 1,
    }),
  ).toBe(42);
  const backup = makeSnapshot(
    engine,
    {
      kind: "pane",
      pane: {
        id: "cache-test",
        workbookName: "Cache test",
        sheetName: "Sheet1",
        cell: "B1",
      },
    },
    "cache-test",
  );
  // Simulate a warm backup with stale derived results, leaving source data intact.
  // This tests the real public API through the UI, without a hidden engine hook.
  const native = JSON.parse(backup.engine);
  let poisoned = 0;
  for (const node of native.managers.dependency.nodes) {
    if (
      node.evaluationResult?.type === "value" &&
      node.evaluationResult.result?.value === 42
    ) {
      node.evaluationResult.result.value = 999;
      poisoned++;
    }
  }
  expect(poisoned).toBeGreaterThan(0);
  backup.engine = JSON.stringify(native);
  await page.goto("./");
  await expect(gridCell(page, "Launch budget, Overview", "D12")).toBeVisible();
  await page.getByLabel("Restore workspace file").setInputFiles({
    name: "stale.workspace.json",
    mimeType: "application/json",
    buffer: Buffer.from(JSON.stringify(backup)),
  });
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  const result = gridCell(page, "Cache test, Sheet1", "B1");
  await expect(result).toHaveText("999");
  await enableDebug(page);
  await clearCache(page);
  await expect(result).toHaveText("42");
  await expect(formulaBar(page)).toHaveValue("=A1*2");
  await expect(gridCell(page, "Cache test, Sheet1", "A1")).toHaveText("21");
  await expect(saved(page)).toBeVisible();
  await page.reload();
  await expect(result).toHaveText("42");
  await gridCell(page, "Cache test, Sheet1", "A1").click();
  await formulaBar(page).fill("10");
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
  await expect(result).toHaveText("20");
});

test("debug tools remain accessible on mobile and can be disabled", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("./");
  await enableDebug(page);
  const panel = page.getByRole("region", { name: "Debug tools" });
  const clear = page.getByRole("button", {
    name: "Clear formula cache",
    exact: true,
  });
  await expect(clear).toBeInViewport();
  await clearCache(page);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(
    390,
  );
  await page.getByRole("button", { name: "Disable debug mode" }).click();
  await expect(panel).toHaveCount(0);
  await enableDebug(page);
  await page.getByRole("button", { name: "Show sidebar", exact: true }).click();
  await page.getByRole("button", { name: "Workspace", exact: true }).click();
  const toggle = page.getByRole("menuitemcheckbox", {
    name: "Debug mode",
    exact: true,
  });
  await expect(toggle).toHaveAttribute("aria-checked", "true");
  await toggle.click();
  await expect(panel).toHaveCount(0);
});
