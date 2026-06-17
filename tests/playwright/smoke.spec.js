// @ts-check
const { test, expect } = require("@playwright/test");

test.describe("Banki import konvertáló - E2E smoke", () => {
  test("loads home and renders all major regions", async ({ page }) => {
    await page.goto("/");
    await expect(page).toHaveTitle(/Banki import konvertáló/);
    await expect(page.locator("#convertBtn")).toBeVisible();
    await expect(page.locator("#registryPill")).toBeVisible();
    await expect(page.locator("#previewBox")).toBeVisible();
  });

  test("opens the Import dialog and shows bank/format selectors", async ({ page }) => {
    await page.goto("/");
    await page.click("#openImportBtn");
    await expect(page.locator("#importDialog")).toBeVisible();
    await expect(page.locator("#bankSelect")).toBeVisible();
    await expect(page.locator("#formatSelect")).toBeVisible();
  });

  test("opens Partners dialog and shows the filter input", async ({ page }) => {
    await page.goto("/");
    await page.click("#partnersBtn");
    await expect(page.locator("#partnersDialog")).toBeVisible();
    // filter.js injects a search input
    await expect(page.locator("#partnersDialog input[type='search'], #partnersDialog input[placeholder*='Keres']")).toHaveCount(1);
  });

  test("MNB pill is keyboard-activatable", async ({ page }) => {
    await page.goto("/");
    const pill = page.locator("#registryPill");
    await pill.focus();
    await expect(pill).toHaveAttribute("role", "button");
    await expect(pill).toHaveAttribute("tabindex", "0");
  });

  test("theme toggle switches dark mode", async ({ page }) => {
    await page.goto("/");
    const html = page.locator("html");
    const before = await html.getAttribute("class");
    const toggle = page.locator("#themeToggle, [data-theme-toggle]");
    if (await toggle.count()) {
      await toggle.first().click();
      const after = await html.getAttribute("class");
      expect(after).not.toEqual(before);
    }
  });

  test("Summary button appears next to Convert", async ({ page }) => {
    await page.goto("/");
    await expect(page.locator("#summaryBtn")).toBeVisible();
  });

  test("static assets respond with 200", async ({ request }) => {
    for (const path of [
      "/static/styles-base.css",
      "/static/tokens.css",
      "/static/app.js",
      "/static/js/toast.js",
      "/static/js/export-summary.js",
      "/static/js/diff-view.js",
    ]) {
      const r = await request.get(path);
      expect(r.status(), path).toBe(200);
    }
  });

  test("path traversal on /static is blocked", async ({ request }) => {
    const r = await request.get("/static/../app.py");
    expect([400, 403, 404]).toContain(r.status());
  });
});
