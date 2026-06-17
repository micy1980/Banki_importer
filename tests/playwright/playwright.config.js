// @ts-check
const { defineConfig } = require("@playwright/test");
module.exports = defineConfig({
  testDir: ".",
  timeout: 30_000,
  use: {
    baseURL: process.env.BANKI_URL || "http://127.0.0.1:8765",
    headless: true,
    viewport: { width: 1280, height: 800 },
  },
  reporter: [["list"]],
});
