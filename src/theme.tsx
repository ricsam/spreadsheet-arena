import { useSyncExternalStore } from "react";

type Theme = "light" | "dark";
const key = "spreadsheet-arena-theme";
const listeners = new Set<() => void>();
let preference: Theme | null = null;
let current: Theme = "light";
let initialized = false;
function readPreference(): Theme | null {
  try {
    const value = localStorage.getItem(key);
    return value === "light" || value === "dark" ? value : null;
  } catch {
    return null;
  }
}
function apply() {
  current =
    preference ??
    (window.matchMedia("(prefers-color-scheme: dark)").matches
      ? "dark"
      : "light");
  document.documentElement.classList.toggle("dark", current === "dark");
  document.documentElement.style.colorScheme = current;
  listeners.forEach((listener) => listener());
}
export function initializeTheme() {
  if (initialized || typeof window === "undefined") return;
  initialized = true;
  preference = readPreference();
  window
    .matchMedia("(prefers-color-scheme: dark)")
    .addEventListener("change", apply);
  window.addEventListener("storage", (event) => {
    if (event.key === key || event.key === null) {
      preference = readPreference();
      apply();
    }
  });
  apply();
}
function subscribe(listener: () => void) {
  initializeTheme();
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}
function toggleTheme() {
  initializeTheme();
  preference = current === "dark" ? "light" : "dark";
  try {
    localStorage.setItem(key, preference);
  } catch {
    /* Remain usable when storage is unavailable. */
  }
  apply();
}
export function useTheme(): { theme: Theme; toggleTheme: () => void } {
  const theme = useSyncExternalStore(
    subscribe,
    () => current,
    () => "light" as Theme,
  );
  return { theme, toggleTheme };
}
