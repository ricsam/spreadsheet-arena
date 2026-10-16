import { test, expect } from "@playwright/test";
import { save } from "./save";

test("formula-first shell, resizable sidebars, and persistent dark theme", async ({
  page,
}) => {
  await page.goto("./");
  const library = page.getByRole("complementary", { name: "Workbook library" });
  const inspector = page.getByRole("complementary", {
    name: "Workbook inspector",
  });
  await expect(
    library.getByRole("heading", { name: "spreadsheet arena" }),
  ).toBeVisible();
  await expect(
    page.locator(".app-header, .workspace-title, .workspace-subtitle"),
  ).toHaveCount(0);
  await expect(
    library.getByRole("button", { name: "Import", exact: true }),
  ).toBeVisible();
  await expect(
    inspector.getByRole("tab", { name: "Cell", exact: true }),
  ).toBeVisible();
  const left = page.getByRole("separator", { name: "Resize Workbook library" });
  const right = page.getByRole("separator", {
    name: "Resize Workbook inspector",
  });
  const before = Number(await left.getAttribute("aria-valuenow"));
  await left.focus();
  await left.press("ArrowRight");
  await expect(left).toHaveAttribute("aria-valuenow", String(before + 8));
  const rightBefore = Number(await right.getAttribute("aria-valuenow"));
  const handle = (await right.boundingBox())!;
  await page.mouse.move(handle.x + handle.width / 2, handle.y + 200);
  await page.mouse.down();
  await page.mouse.move(handle.x + handle.width / 2 - 35, handle.y + 200);
  await page.mouse.up();
  await expect(right).toHaveAttribute(
    "aria-valuenow",
    String(rightBefore + 35),
  );
  const styledCell = page
    .getByRole("grid", { name: "Launch budget, Overview", exact: true })
    .locator('[data-cell-address="A1"]');
  const originalFill = await styledCell.evaluate(
    (el) => getComputedStyle(el).backgroundColor,
  );
  const formula = page.getByRole("textbox", {
    name: "Cell value or formula",
    exact: true,
  });
  await formula.fill("Theme switch draft");
  const cellHandle = await styledCell.elementHandle();
  await page.getByRole("button", { name: "Switch to dark mode" }).click();
  await expect(page.locator("html")).toHaveClass(/dark/);
  await expect(formula).toHaveValue("Theme switch draft");
  expect(await cellHandle!.evaluate((el) => el.isConnected)).toBe(true);
  await expect
    .poll(() =>
      page
        .locator("canvas")
        .first()
        .evaluate(
          (el) => (el as HTMLCanvasElement).getContext("2d")!.strokeStyle,
        ),
    )
    .toBe("#343b44");
  await formula.press("Escape");
  await expect(styledCell).toHaveCSS("background-color", originalFill);
  await expect(page.locator(".formula-bar")).toHaveCSS(
    "background-color",
    "rgb(23, 26, 30)",
  );
  await save(page);
  await page.reload();
  await expect(page.locator("html")).toHaveClass(/dark/);
  await expect(left).toHaveAttribute("aria-valuenow", String(before + 8));
  await expect(right).toHaveAttribute(
    "aria-valuenow",
    String(rightBefore + 35),
  );
  await page
    .getByRole("button", { name: "Hide inspector", exact: true })
    .click();
  await expect(inspector).toHaveCount(0);
  await page
    .getByRole("button", { name: "Show inspector", exact: true })
    .click();
  await expect(inspector).toBeVisible();
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(library).toHaveCount(0);
  await expect(inspector).toHaveCount(0);
  await expect(
    page.getByRole("button", { name: "Show sidebar", exact: true }),
  ).toBeInViewport();
});

test("mobile sidebar controls overlay without overflowing the workspace", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("./");
  await page.getByRole("button", { name: "Show sidebar", exact: true }).click();
  await expect(
    page.getByRole("button", { name: "Import", exact: true }),
  ).toBeInViewport();
  await expect(
    page.getByRole("button", { name: "Workspace", exact: true }),
  ).toBeInViewport();
  await page.getByRole("button", { name: "Hide sidebar", exact: true }).click();
  await page
    .getByRole("button", { name: "Show inspector", exact: true })
    .click();
  await expect(
    page.getByRole("tab", { name: "Cell", exact: true }),
  ).toBeInViewport();
  await page
    .getByRole("button", { name: "Close inspector", exact: true })
    .click();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(
    390,
  );
});
