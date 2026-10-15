import { test, expect } from "@playwright/test";
import { readdir } from "node:fs/promises";

test("production bundles serve separate source maps with source content", async ({
  request,
}) => {
  const scripts = (await readdir("dist/assets")).filter((name) =>
    name.endsWith(".js"),
  );
  expect(scripts.length).toBeGreaterThan(0);
  let appSourceFound = false;
  for (const name of scripts) {
    const response = await request.get(`assets/${name}`);
    expect(response.ok()).toBe(true);
    const script = await response.text();
    const sourceMap = /\/\/# sourceMappingURL=(\S+)\s*$/.exec(script)?.[1];
    expect(sourceMap).toBe(`${name}.map`);
    const mapResponse = await request.get(`assets/${sourceMap}`);
    expect(mapResponse.ok()).toBe(true);
    const map = await mapResponse.json();
    expect(map.version).toBe(3);
    expect(map.file).toBe(name);
    expect(map.mappings.length).toBeGreaterThan(0);
    expect(map.sources.length).toBeGreaterThan(0);
    expect(map.sourcesContent).toHaveLength(map.sources.length);
    const appIndex = map.sources.findIndex((source: string) =>
      source.endsWith("src/App.tsx"),
    );
    if (appIndex !== -1) {
      expect(map.sourcesContent[appIndex]).toContain("export function App()");
      appSourceFound = true;
    }
  }
  expect(appSourceFound).toBe(true);
});
