import { expect, type Page } from "@playwright/test";

export async function save(page: Page) {
  await page.getByRole("button", { name: "Save", exact: true }).click();
  await expect(
    page.getByRole("status").filter({ hasText: "Saved on this device" }),
  ).toBeVisible();
}
