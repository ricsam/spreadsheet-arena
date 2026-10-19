import { test, expect } from "@playwright/test";
import { FormulaEngine } from "@ricsam/formula-engine";
import { makeSnapshot } from "../src/model";

test("inline edits preserve unrelated warm results and update only affected formulas", async ({
  page,
}) => {
  const engine = new FormulaEngine();
  engine.addWorkbook({
    workbookName: "Inline cache",
    data: {
      sheets: [
        {
          name: "Data",
          content: {
            A1: 7,
            B1: "=A1*2",
            C1: "=40+2",
            F3: 12345,
          },
        },
      ],
    },
  });
  const address = (colIndex: number) => ({
    workbookName: "Inline cache",
    sheetName: "Data",
    rowIndex: 0,
    colIndex,
  });
  expect(engine.getCellValue(address(1))).toBe(14);
  expect(engine.getCellValue(address(2))).toBe(42);
  const snapshot = makeSnapshot(
    engine,
    {
      kind: "pane",
      pane: {
        id: "inline-cache",
        workbookName: "Inline cache",
        sheetName: "Data",
        cell: "F3",
      },
    },
    "inline-cache",
  );
  // Make cache reuse observable through the actual UI, with no hidden engine API.
  // A global reset would turn this sentinel back into 42 after ANY inline edit.
  const native = JSON.parse(snapshot.engine);
  let changed = 0;
  for (const node of native.managers.dependency.nodes) {
    if (
      node.evaluationResult?.type === "value" &&
      node.evaluationResult.result?.value === 42
    ) {
      node.evaluationResult.result.value = 4242;
      changed++;
    }
  }
  expect(changed).toBeGreaterThan(0);
  snapshot.engine = JSON.stringify(native);
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("./");
  await expect(page.getByRole("main")).toBeVisible();
  await page.getByLabel("Restore workspace file").setInputFiles({
    name: "inline-cache.workspace.json",
    mimeType: "application/json",
    buffer: Buffer.from(JSON.stringify(snapshot)),
  });
  await page.getByRole("button", { name: "Confirm", exact: true }).click();
  const grid = page.getByRole("grid", {
    name: "Inline cache, Data",
    exact: true,
  });
  const cell = (ref: string) => grid.locator(`[data-cell-address="${ref}"]`);
  await expect(cell("C1")).toHaveText("4,242");
  await expect(cell("B1")).toHaveText("14");
  await cell("F3").dblclick();
  const editor = page.getByRole("textbox", { name: "Edit F3", exact: true });
  await editor.fill("12346");
  await editor.press("Enter");
  await expect(cell("F3")).toHaveText("12,346");
  await expect(cell("C1")).toHaveText("4,242");
  await expect(cell("B1")).toHaveText("14");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(cell("F3")).toHaveText("12,345");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(cell("F3")).toHaveText("12,346");
  await cell("A1").dblclick();
  const inputEditor = page.getByRole("textbox", {
    name: "Edit A1",
    exact: true,
  });
  await inputEditor.fill("9");
  await inputEditor.press("Enter");
  await expect(cell("B1")).toHaveText("18");
  await expect(cell("C1")).toHaveText("4,242");
  await cell("F3").dblclick();
  await editor.fill("");
  await editor.press("Enter");
  await expect(cell("F3")).toHaveText("");
  await expect(cell("C1")).toHaveText("4,242");
  await expect(
    page.getByRole("status").filter({ hasText: "Unsaved changes" }),
  ).toBeVisible();
  expect(errors).toEqual([]);
});
