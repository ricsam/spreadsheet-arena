import { test, expect } from "@playwright/test";
import { save } from "./save";

test("workbook filtering has a recoverable no-results state", async ({
  page,
}) => {
  await page.goto("./");
  const library = page.getByRole("complementary", { name: "Workbook library" });
  const filter = library.getByRole("textbox", { name: "Filter workbooks" });
  await filter.fill("  not-a-workbook  ");
  await expect(library.getByRole("status")).toContainText(
    "No matching workbooks",
  );
  await expect(library.getByRole("button", { name: /^Open / })).toHaveCount(0);
  await library
    .getByRole("button", { name: "Clear search", exact: true })
    .click();
  await expect(filter).toHaveValue("");
  await expect(
    library.getByRole("button", { name: "Open Launch budget", exact: true }),
  ).toBeVisible();
  await filter.fill("  launch  ");
  await expect(
    library.getByRole("button", { name: "Open Launch budget", exact: true }),
  ).toBeVisible();
  await library.getByRole("button", { name: "Clear workbook filter" }).click();
  await expect(
    library.getByRole("button", { name: "Open Assumptions", exact: true }),
  ).toBeVisible();
});

test("narrow screens focus one pane without losing the desktop split layout", async ({
  page,
}) => {
  await page.goto("./");
  const divider = page.getByRole("separator", {
    name: "Resize columns",
    exact: true,
  });
  await divider.focus();
  await divider.press("ArrowLeft");
  const ratio = await divider.getAttribute("aria-valuenow");
  const paneIds = await page
    .locator(".workbook-pane")
    .evaluateAll((els) => els.map((el) => el.getAttribute("data-pane-id")));
  const budget = page.getByRole("grid", {
    name: "Launch budget, Overview",
    exact: true,
  });
  const assumptions = page.getByRole("grid", {
    name: "Assumptions, Inputs",
    exact: true,
  });
  const budgetHandle = await budget.elementHandle();
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(budget).toBeVisible();
  await expect(assumptions).toBeHidden();
  expect(await budgetHandle!.evaluate((el) => el.isConnected)).toBe(true);
  const switcher = page.getByRole("combobox", {
    name: "Open pane",
    exact: true,
  });
  await switcher.selectOption(paneIds[1]!);
  await expect(assumptions).toBeVisible();
  await expect(budget).toBeHidden();
  await page
    .getByRole("textbox", { name: "Cell value or formula", exact: true })
    .fill("72");
  await page
    .getByRole("button", { name: "Apply cell value", exact: true })
    .click();
  await expect(assumptions.locator('[data-cell-address="B4"]')).toHaveText(
    "72",
  );
  expect(
    await page
      .locator(".tile-stage")
      .evaluate((el) => el.scrollWidth <= el.clientWidth),
  ).toBe(true);
  await page.getByRole("button", { name: "Show sidebar", exact: true }).click();
  await page
    .getByRole("button", { name: "Open Launch budget", exact: true })
    .click();
  await expect(
    page.getByRole("complementary", { name: "Workbook library" }),
  ).toHaveCount(0);
  await expect(budget).toBeVisible();
  await page.setViewportSize({ width: 1440, height: 920 });
  await expect(divider).toHaveAttribute("aria-valuenow", ratio!);
  await expect(budget).toBeVisible();
  await expect(assumptions).toBeVisible();
  expect(
    await page
      .locator(".workbook-pane")
      .evaluateAll((els) => els.map((el) => el.getAttribute("data-pane-id"))),
  ).toEqual(paneIds);
  await expect(assumptions.locator('[data-cell-address="B4"]')).toHaveText(
    "72",
  );
});

test("sheet tabs use arrow and Home/End navigation with a single tab stop", async ({
  page,
}) => {
  await page.goto("./");
  await page
    .getByRole("button", { name: "Add sheet to Launch budget", exact: true })
    .click();
  const tabs = page.getByRole("tablist", {
    name: "Sheets in Launch budget",
    exact: true,
  });
  const first = tabs.getByRole("tab").first();
  const last = tabs.getByRole("tab").last();
  await expect(last).toHaveAttribute("aria-selected", "true");
  await last.focus();
  await last.press("ArrowLeft");
  await expect(first).toBeFocused();
  await expect(first).toHaveAttribute("aria-selected", "true");
  await expect(last).toHaveAttribute("tabindex", "-1");
  await first.press("End");
  await expect(last).toBeFocused();
  await last.press("ArrowRight");
  await expect(first).toBeFocused();
  await first.press("End");
  await last.press("Home");
  await expect(first).toBeFocused();
  await expect(
    page.locator(`#${await first.getAttribute("aria-controls")}`),
  ).toBeVisible();
});

test("formula drafts have an explicit cancel action without changing the workbook", async ({
  page,
}) => {
  await page.goto("./");
  const formula = page.getByRole("textbox", {
    name: "Cell value or formula",
    exact: true,
  });
  await expect(formula).toHaveValue("=SUM(D9:D10)");
  await save(page);
  await formula.fill("discard this draft");
  await expect(page.locator(".formula-bar")).toHaveClass(/is-editing/);
  await page
    .getByRole("button", { name: "Cancel cell edit", exact: true })
    .click();
  await expect(formula).toHaveValue("=SUM(D9:D10)");
  await expect(formula).toBeFocused();
  await expect(page.locator(".formula-bar")).not.toHaveClass(/is-editing/);
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
});

test("closing all panes explains that workbooks can be reopened", async ({
  page,
}) => {
  await page.goto("./");
  await page
    .getByRole("button", { name: "Close Launch budget pane", exact: true })
    .click();
  await page
    .getByRole("button", { name: "Close Assumptions pane", exact: true })
    .click();
  await expect(
    page.getByRole("heading", { name: "Your workbooks are still here" }),
  ).toBeVisible();
  await page
    .getByRole("button", { name: "Open workbook library", exact: true })
    .click();
  await page
    .getByRole("button", { name: "Open Launch budget", exact: true })
    .click();
  await expect(
    page.getByRole("grid", { name: "Launch budget, Overview", exact: true }),
  ).toBeVisible();
});
