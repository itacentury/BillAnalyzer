import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // dom.js touches window.matchMedia at import time, so even the "pure"
    // function tests need a DOM environment. happy-dom implements matchMedia
    // (jsdom does not without a manual stub).
    environment: "happy-dom",
    include: ["tests/frontend/**/*.test.js"],
  },
});
