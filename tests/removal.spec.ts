import { test, expect, type Page } from "@playwright/test";
import { FormulaEngine } from "@ricsam/formula-engine";
import { makeSnapshot } from "../src/model";
import { save } from "./save";

const book = (page: Page, name: string) =>
  page.getByRole("button", { name: `Open ${name}`, exact: true });
const grid = (page: Page, name: string) =>
  page.getByRole("grid", { name, exact: true });
const formulaBar = (page: Page) =>
  page.getByRole("textbox", { name: "Cell value or formula", exact: true });
async function remove(page: Page, name: string) {
  await page
    .getByRole("button", { name: `Remove ${name}`, exact: true })
    .click();
  await page
    .getByRole("button", { name: "Remove workbook", exact: true })
    .click();
  await expect(book(page, name)).toHaveCount(0);
  await expect(page.getByRole("dialog")).toHaveCount(0);
}

test("remove the last imported workbook, persist an empty workspace, and create again", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  // Synthetic data only; reproduces the reported name, sheet count and active sheet.
  const name = "16.5.82.1_Responder_Peptide_Pulse_Workbook";
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: name,
    data: {
      sheets: Array.from({ length: 13 }, (_, index) => ({
        name: `Sheet${index + 1}`,
        content: { A1: index, B1: "=A1*2" },
      })),
    },
  });
  const backup = makeSnapshot(
    engine,
    {
      kind: "pane",
      pane: {
        id: "imported",
        workbookName: name,
        sheetName: "Sheet13",
        cell: "B1",
      },
    },
    "imported",
  );
  await page.goto("./");
  await expect(grid(page, "Launch budget, Overview")).toBeVisible();
  await page.getByLabel("Restore workspace file").setInputFiles({
    name: "synthetic.workspace.json",
    mimeType: "application/json",
    buffer: Buffer.from(JSON.stringify(backup)),
  });
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  await expect(formulaBar(page)).toHaveValue("=A1*2");
  await page
    .getByRole("button", { name: `Remove ${name}`, exact: true })
    .click();
  await page.getByRole("button", { name: "Cancel", exact: true }).click();
  await expect(book(page, name)).toBeVisible();
  await remove(page, name);
  await expect(page.getByRole("grid")).toHaveCount(0);
  await expect(formulaBar(page)).toBeDisabled();
  await expect(formulaBar(page)).toHaveValue("");
  await expect(
    page.getByRole("button", { name: "Split active pane right" }),
  ).toBeDisabled();
  await save(page);
  await page.reload();
  await expect(
    page.getByRole("button", { name: "Create a workbook", exact: true }),
  ).toBeVisible();
  await expect(
    page.getByRole("navigation", { name: "Workbooks" }).getByRole("button"),
  ).toHaveCount(0);
  await page
    .getByRole("button", { name: "Create a workbook", exact: true })
    .click();
  await expect(grid(page, "Untitled workbook, Sheet1")).toBeVisible();
  expect(errors).toEqual([]);
});

test("remove active maximized duplicate panes, undo and redo without losing other workbooks", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("./");
  const total = grid(page, "Launch budget, Overview").locator(
    '[data-cell-address="D12"]',
  );
  await expect(total).toHaveText("7,535");
  await grid(page, "Assumptions, Inputs")
    .locator('[data-cell-address="B4"]')
    .click();
  await page
    .getByRole("button", { name: "Split Assumptions below", exact: true })
    .click();
  await page
    .getByRole("button", { name: "Maximize Assumptions", exact: true })
    .last()
    .click();
  await remove(page, "Assumptions");
  await expect(page.getByRole("grid")).toHaveCount(1);
  await expect(grid(page, "Launch budget, Overview")).toBeVisible();
  await expect(formulaBar(page)).toHaveValue("=SUM(D9:D10)");
  await expect(
    page.getByRole("button", { name: "Restore panes", exact: true }),
  ).toHaveCount(0);
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(book(page, "Assumptions")).toBeVisible();
  await book(page, "Assumptions").click();
  await expect(
    grid(page, "Assumptions, Inputs").locator('[data-cell-address="B4"]'),
  ).toHaveText("45");
  await expect(total).toHaveText("7,535");
  await page
    .getByRole("button", { name: "Split Assumptions right", exact: true })
    .click();
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(book(page, "Assumptions")).toHaveCount(0);
  await expect(page.getByRole("grid")).toHaveCount(1);
  await save(page);
  await page.reload();
  await expect(grid(page, "Launch budget, Overview")).toBeVisible();
  await expect(book(page, "Assumptions")).toHaveCount(0);
  expect(errors).toEqual([]);
});

test("remove an inactive open workbook without disrupting the active pane", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("./");
  await grid(page, "Assumptions, Inputs")
    .locator('[data-cell-address="B4"]')
    .click();
  await remove(page, "Launch budget");
  await expect(grid(page, "Assumptions, Inputs")).toBeVisible();
  await expect(formulaBar(page)).toHaveValue("45");
  await save(page);
  await page.reload();
  await expect(formulaBar(page)).toHaveValue("45");
  await expect(book(page, "Launch budget")).toHaveCount(0);
  expect(errors).toEqual([]);
});

test("undoing sheet and workbook creation reconciles panes before rendering", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("./");
  await page.getByRole("button", { name: "New workbook", exact: true }).click();
  await page.getByRole("textbox", { name: "Workbook name" }).fill("Temporary");
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  await page
    .getByRole("button", { name: "Add sheet to Temporary", exact: true })
    .click();
  await expect(grid(page, "Temporary, Sheet2")).toBeVisible();
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(grid(page, "Temporary, Sheet1")).toBeVisible();
  await expect(
    page.getByRole("textbox", { name: "Cell address", exact: true }),
  ).toHaveValue("A1");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(book(page, "Temporary")).toHaveCount(0);
  await expect(page.getByRole("grid")).toHaveCount(2);
  await save(page);
  expect(errors).toEqual([]);
});
