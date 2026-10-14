import { defineConfig } from "@playwright/test";
export default defineConfig({
  testDir: "./tests",
  fullyParallel: false,
  workers: 1,
  use: {
    baseURL: "http://127.0.0.1:5178/spreadsheet-arena/",
    viewport: { width: 1440, height: 920 },
    trace: "retain-on-failure",
  },
  webServer: {
    command: "npm run preview -- --port 5178 --strictPort",
    url: "http://127.0.0.1:5178/spreadsheet-arena/",
    reuseExistingServer: !process.env.CI,
  },
});
