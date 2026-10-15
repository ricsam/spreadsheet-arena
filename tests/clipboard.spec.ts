import { test, expect, type Page } from "@playwright/test";
const cell = (page: Page, grid: string, address: string) =>
  page
    .getByRole("grid", { name: grid, exact: true })
    .locator(`[data-cell-address="${address}"]`);
const quick = (page: Page, name: string) =>
  page
    .getByRole("group", { name: "Quick cell formatting" })
    .first()
    .getByRole("button", { name, exact: true });

test("cross-workbook toolbar paste keeps formatting, plain paste keeps destination formatting, denied OS read falls back", async ({
  page,
  context,
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.goto("./");
  const source = cell(page, "Assumptions, Inputs", "B4");
  const target = cell(page, "Launch budget, Overview", "B5");
  await source.click();
  await quick(page, "Italic").click();
  await page.getByRole("button", { name: "Copy cells", exact: true }).click();
  await target.click();
  await page.getByRole("button", { name: "Paste", exact: true }).click();
  await expect(target).toHaveText("45");
  await expect(quick(page, "Italic")).toHaveAttribute("aria-pressed", "true");
  await page.getByRole("button", { name: "Undo", exact: true }).click();
  await expect(target).toHaveText("80");
  await expect(quick(page, "Italic")).toHaveAttribute("aria-pressed", "false");
  await page
    .getByRole("button", { name: "Paste options", exact: true })
    .click();
  await page
    .getByRole("menuitem", { name: "Paste without formatting", exact: true })
    .click();
  await expect(target).toHaveText("45");
  await expect(quick(page, "Italic")).toHaveAttribute("aria-pressed", "false");
  // A new OS clipboard value takes precedence over our old internal snapshot.
  await page.evaluate(() => navigator.clipboard.writeText("123"));
  await page.getByRole("button", { name: "Paste", exact: true }).click();
  await expect(target).toHaveText("123");
  await page.evaluate(() => {
    Object.defineProperty(navigator.clipboard, "readText", {
      configurable: true,
      value: () =>
        Promise.reject(new DOMException("Denied", "NotAllowedError")),
    });
  });
  await page.getByRole("button", { name: "Paste", exact: true }).click();
  await expect(target).toHaveText("45");
});

test("keyboard copy/paste remains global and text fields keep normal paste", async ({
  page,
  context,
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await page.goto("./");
  await cell(page, "Assumptions, Inputs", "B4").click();
  await page.keyboard.press("ControlOrMeta+c");
  const target = cell(page, "Launch budget, Overview", "B5");
  await target.click();
  await page.keyboard.press("ControlOrMeta+v");
  await expect(target).toHaveText("45");
  await page.evaluate(() => navigator.clipboard.writeText("external text"));
  const editor = page.getByRole("textbox", {
    name: "Cell value or formula",
    exact: true,
  });
  await editor.focus();
  await editor.press("ControlOrMeta+a");
  await editor.press("ControlOrMeta+v");
  await expect(editor).toHaveValue("external text");
  await expect(target).toHaveText("45");
});
