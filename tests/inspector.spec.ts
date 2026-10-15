import { test, expect } from "@playwright/test";

test("range formatting survives toolbar and tab focus; numeric drafts retain their target", async ({
  page,
}) => {
  await page.goto("./");
  const grid = page.getByRole("grid", {
    name: "Launch budget, Overview",
    exact: true,
  });
  await expect(grid).toBeVisible();
  const cell = (name: string) => grid.locator(`[data-cell-address="${name}"]`);
  const inspector = page.getByRole("complementary", {
    name: "Workbook inspector",
  });
  if (!(await inspector.isVisible()))
    await page
      .getByRole("button", { name: "Show inspector", exact: true })
      .click();
  await cell("B5").click();
  await cell("C6").click({ modifiers: ["Shift"] });
  await inspector.getByRole("tab", { name: "Sheet", exact: true }).click();
  await inspector.getByRole("tab", { name: "Cell", exact: true }).click();
  // First group is the formula-bar quick action group, not the inspector group.
  await page
    .getByRole("group", { name: "Quick cell formatting" })
    .first()
    .getByRole("button", { name: "Italic", exact: true })
    .click();
  for (const name of ["B5", "C5", "B6", "C6"]) {
    await cell(name).click();
    await expect(
      inspector.getByRole("button", { name: "Italic", exact: true }),
    ).toHaveAttribute("aria-pressed", "true");
  }
  await cell("D6").click();
  await expect(
    inspector.getByRole("button", { name: "Italic", exact: true }),
  ).toHaveAttribute("aria-pressed", "false");
  await cell("B5").click();
  await inspector
    .getByRole("spinbutton", { name: "Font size", exact: true })
    .fill("22");
  await cell("D6").click();
  await expect(
    inspector.getByRole("spinbutton", { name: "Font size", exact: true }),
  ).toHaveValue("11");
  await cell("B5").click();
  await expect(
    inspector.getByRole("spinbutton", { name: "Font size", exact: true }),
  ).toHaveValue("22");
  await inspector.getByRole("tab", { name: "Sheet", exact: true }).click();
  await inspector
    .getByRole("spinbutton", { name: "Row height", exact: true })
    .fill("44");
  await inspector
    .getByRole("spinbutton", { name: "Row height", exact: true })
    .press("Enter");
  await expect(
    inspector.getByRole("spinbutton", { name: "Row height", exact: true }),
  ).toHaveValue("44");
  await cell("B6").click();
  await expect(
    inspector.getByRole("spinbutton", { name: "Row height", exact: true }),
  ).not.toHaveValue("44");
});

test("inspector formatting, dimensions, and table definitions undo", async ({
  page,
}) => {
  await page.goto("./");
  const grid = page.getByRole("grid", {
    name: "Launch budget, Overview",
    exact: true,
  });
  await expect(grid).toBeVisible();
  await grid.locator('[data-cell-address="B5"]').click();
  const inspector = page.getByRole("complementary", {
    name: "Workbook inspector",
  });
  if (!(await inspector.isVisible()))
    await page
      .getByRole("button", { name: "Show inspector", exact: true })
      .click();
  await inspector.getByRole("tab", { name: "Cell", exact: true }).click();
  const bold = inspector.getByRole("button", { name: "Bold", exact: true });
  await bold.click();
  await expect(bold).toHaveAttribute("aria-pressed", "true");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(bold).toHaveAttribute("aria-pressed", "false");
  await inspector.getByRole("tab", { name: "Sheet", exact: true }).click();
  await inspector
    .getByRole("spinbutton", { name: "Column width", exact: true })
    .fill("155");
  await inspector
    .getByRole("spinbutton", { name: "Column width", exact: true })
    .press("Enter");
  await expect(
    inspector.getByRole("spinbutton", { name: "Column width", exact: true }),
  ).toHaveValue("155");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(
    inspector.getByRole("spinbutton", { name: "Column width", exact: true }),
  ).toHaveValue("80");
  await inspector.getByRole("tab", { name: "Table", exact: true }).click();
  await inspector
    .getByRole("textbox", { name: "Table name", exact: true })
    .fill("InspectorItems");
  await inspector
    .getByRole("button", { name: "Create from selection", exact: true })
    .click();
  await expect(
    inspector.getByRole("textbox", { name: "Table range", exact: true }),
  ).toHaveValue("B5:B5");
  await inspector
    .getByRole("textbox", { name: "Table name", exact: true })
    .fill("RenamedItems");
  await inspector
    .getByRole("button", { name: "Rename table", exact: true })
    .click();
  await expect(
    inspector.getByRole("combobox", { name: "Table", exact: true }),
  ).toContainText("RenamedItems");
  await grid.locator('[data-cell-address="B5"]').click();
  await grid
    .locator('[data-cell-address="C6"]')
    .click({ modifiers: ["Shift"] });
  await inspector
    .getByRole("button", { name: "Resize to selection", exact: true })
    .click();
  await expect(
    inspector.getByRole("textbox", { name: "Table range", exact: true }),
  ).toHaveValue("B5:C6");
  await inspector
    .getByRole("button", { name: "Remove table definition", exact: true })
    .click();
  await expect(grid.locator('[data-cell-address="B5"]')).toHaveText("80");
  await expect(grid.locator('[data-cell-address="C5"]')).toHaveText("45");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(
    inspector.getByRole("combobox", { name: "Table", exact: true }),
  ).toContainText("RenamedItems");
});
