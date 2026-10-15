import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig({
  base: "/spreadsheet-arena/",
  plugins: [react(), tailwindcss()],
  resolve: { alias: { "@": new URL("./src", import.meta.url).pathname } },
  build: {
    // Public external maps let production DevTools resolve minified traces.
    sourcemap: true,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes("/@ricsam/formula-engine/")) return "formula-engine";
          if (
            id.includes("/@ricsam/selection-manager/") ||
            id.includes("/@ricsam/react-spreadsheets/")
          )
            return "spreadsheet";
          if (
            id.includes("/node_modules/react/") ||
            id.includes("/node_modules/react-dom/") ||
            id.includes("/node_modules/scheduler/")
          )
            return "react";
        },
      },
    },
  },
  test: { include: ["src/**/*.test.ts"], environment: "node" },
});
