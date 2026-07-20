import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Pin the timezone here (not only in the npm script's cross-env) so a bare
    // `npx vitest` is deterministic too. An explicitly set TZ still wins.
    env: { TZ: process.env.TZ ?? "Europe/Berlin" },
    // dom.js touches window.matchMedia at import time, so even the "pure"
    // function tests need a DOM environment. happy-dom implements matchMedia
    // (jsdom does not without a manual stub).
    environment: "happy-dom",
    include: ["tests/frontend/**/*.test.js"],
  },
});
