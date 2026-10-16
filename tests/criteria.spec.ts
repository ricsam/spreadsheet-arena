import { test, expect, type Page } from "@playwright/test";
import { FormulaEngine } from "@ricsam/formula-engine";
import { makeSnapshot } from "../src/model";
import { save } from "./save";

const grid = (page: Page, sheet: string) =>
  page.getByRole("grid", { name: `Blank criteria, ${sheet}`, exact: true });
const cell = (page: Page, sheet: string, address: string) =>
  grid(page, sheet).locator(`[data-cell-address="${address}"]`);

async function expectSummary(
  page: Page,
  sum: string,
  average: string,
  count: string,
) {
  await expect(cell(page, "Summary", "B1")).toHaveText(sum);
  await expect(cell(page, "Summary", "B2")).toHaveText(average);
  await expect(cell(page, "Summary", "B3")).toHaveText(count);
}

test("empty criteria survive sheet switches, edits, undo, cache clearing and reload", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: "Blank criteria",
    data: {
      sheets: [
        {
          name: "Data",
          content: {
            A1: 10,
            A2: 20,
            A3: 30,
            A4: 40,
            A5: 50,
            A6: 60,
            A7: 70,
            A8: 80,
            // Missing cells, a stored blank and a formula-produced blank match;
            // numbers, booleans and nonempty strings must not match "".
            B2: "yes",
            B3: "",
            B4: '=""',
            B5: 0,
            B6: false,
            B7: "no",
          },
        },
        {
          name: "Summary",
          content: {
            A1: "Sum",
            B1: '=SUMIFS(Data!A1:A8,Data!B1:B8,"")',
            A2: "Average",
            B2: '=AVERAGEIFS(Data!A1:A8,Data!B1:B8,"")',
            A3: "Count",
            B3: '=COUNTIFS(Data!B1:B8,"")',
          },
        },
      ],
    },
  });
  const snapshot = makeSnapshot(
    engine,
    {
      kind: "pane",
      pane: {
        id: "blank-criteria",
        workbookName: "Blank criteria",
        sheetName: "Summary",
        cell: "A1",
      },
    },
    "blank-criteria",
  );
  await page.goto("./");
  await expect(
    page.getByRole("grid", { name: "Launch budget, Overview", exact: true }),
  ).toBeVisible();
  await page.getByLabel("Restore workspace file").setInputFiles({
    name: "blank-criteria.workspace.json",
    mimeType: "application/json",
    buffer: Buffer.from(JSON.stringify(snapshot)),
  });
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  await expectSummary(page, "160", "40", "4");

  await page.getByRole("tab", { name: "Data", exact: true }).click();
  await cell(page, "Data", "B1").click();
  await page
    .getByRole("textbox", { name: "Cell value or formula", exact: true })
    .fill("yes");
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
  await page.getByRole("tab", { name: "Summary", exact: true }).click();
  await expectSummary(page, "150", "50", "3");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expectSummary(page, "160", "40", "4");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expectSummary(page, "150", "50", "3");

  await page.getByRole("button", { name: "Workspace", exact: true }).click();
  await page
    .getByRole("menuitemcheckbox", { name: "Debug mode", exact: true })
    .click();
  await page
    .getByRole("button", { name: "Clear formula cache", exact: true })
    .click();
  await expect(
    page.getByRole("status").filter({ hasText: "Formula cache cleared." }),
  ).toBeVisible();
  await expectSummary(page, "150", "50", "3");
  await page.getByLabel("Save contents", { exact: true }).selectOption("warm");
  await save(page);
  await page.reload();
  await expectSummary(page, "150", "50", "3");

  // A restored range cache must also invalidate when a nonempty cell becomes missing.
  await page.getByRole("tab", { name: "Data", exact: true }).click();
  await cell(page, "Data", "B1").click();
  await page
    .getByRole("textbox", { name: "Cell value or formula", exact: true })
    .fill("");
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
  await page.getByRole("tab", { name: "Summary", exact: true }).click();
  await expectSummary(page, "160", "40", "4");
  expect(errors).toEqual([]);
});
