import { test, expect, type Page } from "@playwright/test";

const GRID = "Launch budget, Expenses";

async function openExpenses(page: Page) {
  await page.goto("./");
  await page
    .getByRole("tab", { name: "Expenses", exact: true })
    .first()
    .click();
  const grid = page.getByRole("grid", { name: GRID, exact: true });
  await expect(grid).toBeVisible();
  const cell = (address: string) =>
    grid.locator(`[data-cell-address="${address}"]`);
  const menuButton = (column: string) =>
    grid.getByRole("button", {
      name: new RegExp(`^Sort and filter ${column}`),
    });
  const menu = (column: string) =>
    page.getByRole("dialog", {
      name: `Sort and filter ${column}`,
      exact: true,
    });
  const rowNumbers = async () =>
    (await grid.getByRole("rowheader").allInnerTexts())
      .map(Number)
      .filter((row) => row <= 16);
  return { grid, cell, menuButton, menu, rowNumbers };
}

async function showTableInspector(page: Page) {
  const inspector = page.getByRole("complementary", {
    name: "Workbook inspector",
  });
  if (!(await inspector.isVisible()))
    await page
      .getByRole("button", { name: "Show inspector", exact: true })
      .click();
  await inspector.getByRole("tab", { name: "Table", exact: true }).click();
  return inspector.getByRole("region", { name: "Sort and filter" });
}

test("header menus filter and sort a table; undo, redo and the inspector follow", async ({
  page,
}) => {
  const { cell, menuButton, menu, rowNumbers } = await openExpenses(page);
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));

  await menuButton("Category").click();
  await menu("Category").getByRole("checkbox", { name: "Catering" }).uncheck();
  await menu("Category").getByRole("button", { name: "OK" }).click();
  await expect(menu("Category")).toBeHidden();
  // Rows 3, 8 and 9 hold Catering: hidden, not moved.
  await expect(cell("A3")).toHaveCount(0);
  await expect
    .poll(rowNumbers)
    .toEqual([1, 2, 4, 5, 6, 7, 10, 11, 12, 13, 14, 15, 16]);
  await expect(cell("A4")).toHaveText("Sound system");
  await expect(menuButton("Category")).toHaveAttribute("data-filtered", "true");
  // The total keeps counting every row, like Excel's SUM.
  await expect(cell("D17")).toHaveText("9,335");

  await cell("A2").click();
  const section = await showTableInspector(page);
  await expect(
    section.getByRole("list", { name: "Active filters" }),
  ).toHaveText("Category: Admin, Creative, Equipment, Venue");
  await expect(section).toContainText("3 rows hidden");

  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(cell("A3")).toHaveText("Catering tasting");
  await expect(section).toContainText("0 rows hidden");
  await page.getByRole("button", { name: "Redo", exact: true }).click();
  await expect(cell("A3")).toHaveCount(0);

  await menuButton("Amount").click();
  await menu("Amount")
    .getByRole("button", { name: "Sort largest to smallest" })
    .click();
  // Rows moved: the largest visible amount is on row 3 now, and its formula
  // moved with it. Catering balance (3,420) sorted to row 2, still hidden.
  await expect(cell("A2")).toHaveCount(0);
  await expect(cell("A3")).toHaveText("Venue deposit");
  await expect(cell("F3")).toHaveText("1,500");
  await expect(menuButton("Amount")).toHaveAttribute("data-sort", "descending");
  await expect(section).toContainText("sorted by Amount descending");

  await section.getByRole("button", { name: "Clear filters" }).click();
  await expect(cell("A2")).toHaveText("Catering balance");
  await expect
    .poll(rowNumbers)
    .toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);
  expect(errors).toEqual([]);
});

test("Alt+Down opens a header menu; search filters, Escape cancels", async ({
  page,
}) => {
  const { cell, menu, rowNumbers } = await openExpenses(page);
  await cell("A1").click({ position: { x: 12, y: 10 } });
  await page.keyboard.press("Alt+ArrowDown");
  const items = menu("Item");
  await expect(items).toBeVisible();
  await expect(items.getByRole("searchbox")).toBeFocused();
  await page.keyboard.press("Escape");
  await expect(items).toBeHidden();
  await expect.poll(rowNumbers).toHaveLength(16);

  // Keyboard focus is back on the grid, so Alt+Down opens the menu again.
  await page.keyboard.press("Alt+ArrowDown");
  await items.getByRole("searchbox").fill("venue");
  await expect(items.getByRole("checkbox")).toHaveCount(3);
  await items.getByRole("searchbox").press("Enter");
  await expect(items).toBeHidden();
  await expect.poll(rowNumbers).toEqual([1, 2, 10, 16]);

  // Arrow keys step over the hidden rows.
  await cell("A2").click();
  await page.keyboard.press("ArrowDown");
  await expect(
    page.getByRole("textbox", { name: "Cell address", exact: true }),
  ).toHaveValue("A10");
});

test("copying a filtered range copies the shown rows and pastes them without gaps; Delete clears shown rows only", async ({
  page,
  context,
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  const { cell, menuButton, menu } = await openExpenses(page);
  await menuButton("Category").click();
  const categories = menu("Category");
  await categories.getByRole("checkbox", { name: "(Select all)" }).uncheck();
  await categories.getByRole("checkbox", { name: "Venue" }).check();
  await categories.getByRole("button", { name: "OK" }).click();
  await expect(cell("A3")).toHaveCount(0);

  await cell("A2").click();
  await cell("B13").click({ modifiers: ["Shift"] });
  await page.keyboard.press("ControlOrMeta+c");
  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .toBe("Venue deposit\tVenue\nVenue balance\tVenue\nCleaning crew\tVenue");
  await cell("A20").click();
  await page.keyboard.press("ControlOrMeta+v");
  await expect(cell("A20")).toHaveText("Venue deposit");
  await expect(cell("A21")).toHaveText("Venue balance");
  await expect(cell("B22")).toHaveText("Venue");
  await expect(cell("A23")).toHaveText("");

  await cell("E2").click();
  await cell("E13").click({ modifiers: ["Shift"] });
  await page.keyboard.press("Delete");
  await expect(cell("E2")).toHaveText("");
  await expect(cell("E10")).toHaveText("");
  const section = await showTableInspector(page);
  await section.getByRole("button", { name: "Clear filters" }).click();
  // Rows the filter hid kept their status.
  await expect(cell("E3")).toHaveText("Paid");
  await expect(cell("E4")).toHaveText("Booked");
  await expect(cell("E14")).toHaveText("Booked");
});

test("the fill handle fills only shown rows; clicking it without a drag changes nothing", async ({
  page,
}) => {
  const { grid, cell, menuButton, menu } = await openExpenses(page);
  await menuButton("Category").click();
  const categories = menu("Category");
  await categories.getByRole("checkbox", { name: "(Select all)" }).uncheck();
  await categories.getByRole("checkbox", { name: "Venue" }).check();
  await categories.getByRole("button", { name: "OK" }).click();
  await expect(cell("A3")).toHaveCount(0);

  const handleCenter = async () => {
    const box = (await grid
      .locator('[data-fill-handle="true"]')
      .boundingBox())!;
    return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  };
  // A click on the handle is not a fill.
  await cell("E2").click();
  let handle = await handleCenter();
  await page.mouse.click(handle.x, handle.y);
  await expect(cell("E2")).toHaveText("Paid");

  // Dragging it from row 2 to row 13 fills the shown rows 10 and 13 only.
  handle = await handleCenter();
  const target = (await cell("E13").boundingBox())!;
  await page.mouse.move(handle.x, handle.y);
  await page.mouse.down();
  await page.mouse.move(
    target.x + target.width / 2,
    target.y + target.height / 2,
    {
      steps: 8,
    },
  );
  await page.mouse.up();
  await expect(cell("E10")).toHaveText("Paid");
  await expect(cell("E13")).toHaveText("Paid");

  const section = await showTableInspector(page);
  await section.getByRole("button", { name: "Clear filters" }).click();
  await expect(cell("E4")).toHaveText("Booked");
  await expect(cell("E5")).toHaveText("Quote");
  await expect(cell("E8")).toHaveText("Quote");
  await expect(cell("E10")).toHaveText("Paid");
});
