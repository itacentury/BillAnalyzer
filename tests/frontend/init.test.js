import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

/**
 * Importing app.js *is* the test: module scripts run after parsing, so app.js
 * calls init() at import time unless document.readyState is still "loading" —
 * and it never is under happy-dom. The service-worker block is skipped for free
 * because happy-dom's navigator has no `serviceWorker`.
 *
 * Every feature module is replaced with a spy, so what is under test is purely
 * init()'s control flow: one failing wiring step must not stop the others.
 */

const steps = vi.hoisted(() => {
  const named = (name) => {
    const spy = vi.fn();
    // init() labels a failure with `step.name`, and vi.fn() reports "spy".
    Object.defineProperty(spy, "name", { value: name });
    return spy;
  };
  const names = [
    "setupComboboxes",
    "applyFilter",
    "refreshAllData",
    "setupFilterListeners",
    "setupModalListeners",
    "setupInvoiceListListeners",
    "setupPaginationListeners",
    "setupPageSizeListeners",
    "setupBulkListeners",
    "setupStatsListeners",
    "setupImportListeners",
    "setupCategorizeListeners",
    "initToastListeners",
    "setupKeyboardListeners",
    "setupDrawerListeners",
    "setupSheetGestures",
    "setupViewportListeners",
    "updateFilterBadge",
    "loadInvoices",
  ];
  return Object.fromEntries(names.map((name) => [name, named(name)]));
});

vi.mock("../../static/js/filters.js", () => ({
  applyFilter: steps.applyFilter,
  setupFilterListeners: steps.setupFilterListeners,
  updateFilterBadge: steps.updateFilterBadge,
}));
vi.mock("../../static/js/api.js", () => ({
  loadInvoices: steps.loadInvoices,
  refreshAllData: steps.refreshAllData,
  setupPaginationListeners: steps.setupPaginationListeners,
}));
vi.mock("../../static/js/modals.js", () => ({
  setupModalListeners: steps.setupModalListeners,
}));
vi.mock("../../static/js/render.js", () => ({
  setupInvoiceListListeners: steps.setupInvoiceListListeners,
}));
vi.mock("../../static/js/bulk.js", () => ({
  setupBulkListeners: steps.setupBulkListeners,
}));
vi.mock("../../static/js/stats.js", () => ({
  setupStatsListeners: steps.setupStatsListeners,
}));
vi.mock("../../static/js/import.js", () => ({
  setupImportListeners: steps.setupImportListeners,
}));
vi.mock("../../static/js/categorize.js", () => ({
  setupCategorizeListeners: steps.setupCategorizeListeners,
}));
vi.mock("../../static/js/combobox.js", () => ({
  setupComboboxes: steps.setupComboboxes,
}));
vi.mock("../../static/js/toast.js", () => ({
  initToastListeners: steps.initToastListeners,
}));
vi.mock("../../static/js/keyboard.js", () => ({
  setupKeyboardListeners: steps.setupKeyboardListeners,
}));
vi.mock("../../static/js/drawer.js", () => ({
  setupDrawerListeners: steps.setupDrawerListeners,
}));
vi.mock("../../static/js/sheet.js", () => ({
  setupSheetGestures: steps.setupSheetGestures,
}));
vi.mock("../../static/js/viewport.js", () => ({
  setupViewportListeners: steps.setupViewportListeners,
}));
vi.mock("../../static/js/pagesize.js", () => ({
  setupPageSizeListeners: steps.setupPageSizeListeners,
}));

// Ordered as init() runs them: the pre-load block first, then the wiring loop.
const PRE_LOAD_STEPS = ["setupComboboxes", "applyFilter", "refreshAllData"];
const WIRING_STEPS = [
  "setupFilterListeners",
  "setupModalListeners",
  "setupInvoiceListListeners",
  "setupPaginationListeners",
  "setupPageSizeListeners",
  "setupBulkListeners",
  "setupStatsListeners",
  "setupImportListeners",
  "setupCategorizeListeners",
  "initToastListeners",
  "setupKeyboardListeners",
  "setupDrawerListeners",
  "setupSheetGestures",
  "setupViewportListeners",
];
const ALL_STEPS = [...PRE_LOAD_STEPS, ...WIRING_STEPS];

let consoleError;

/** Import app.js fresh, so init() runs again against the current spy setup. */
const runInit = async () => {
  vi.resetModules();
  await import("../../static/js/app.js");
};

/** Names of the steps that ran, so a failure names the missing wiring. */
const stepsThatRan = (names = ALL_STEPS) =>
  names.filter((name) => steps[name].mock.calls.length > 0);

/** Position of a step's first call in the global invocation sequence. */
const callOrder = (name) => steps[name].mock.invocationCallOrder[0];

/** Make a step throw, the way a missing DOM element would. */
const breakStep = (name) => {
  steps[name].mockImplementation(() => {
    throw new Error(`${name} exploded`);
  });
};

const loggedLabels = () =>
  consoleError.mock.calls.map(([message]) => String(message));

beforeEach(() => {
  for (const step of Object.values(steps)) step.mockReset();
  consoleError = vi.spyOn(console, "error").mockImplementation(() => {});
});

afterEach(() => {
  consoleError.mockRestore();
});

describe("init", () => {
  it("runs every wiring step exactly once", async () => {
    await runInit();
    for (const name of ALL_STEPS) {
      expect(steps[name], name).toHaveBeenCalledTimes(1);
    }
    expect(consoleError).not.toHaveBeenCalled();
  });

  it("instantiates the comboboxes before the first data load", async () => {
    // Load-bearing, not incidental: refreshAllData() fans out to loadStores()
    // and loadCategories(), which feed their options into the combobox
    // instances via getCombobox() — and skip that silently when none exists
    // yet. stepsThatRan() reports the declared order of ALL_STEPS, not the
    // observed one, so only the invocation order catches a reordering.
    await runInit();

    expect(callOrder("setupComboboxes")).toBeLessThan(
      callOrder("refreshAllData"),
    );
  });

  it("keeps wiring the remaining steps when one throws", async () => {
    // The regression itself: setupFilterListeners() threw on a missing,
    // conditionally rendered element and took the sidebar nav, the "+" button,
    // Import, the modals, the keyboard shortcuts and the drawer down with it.
    breakStep("setupFilterListeners");
    await runInit();

    const survivors = WIRING_STEPS.filter(
      (name) => name !== "setupFilterListeners",
    );
    expect(stepsThatRan(survivors)).toEqual(survivors);
    for (const name of survivors) {
      expect(steps[name], name).toHaveBeenCalledTimes(1);
    }
  });

  it("names the failing step in the console", async () => {
    breakStep("setupFilterListeners");
    await runInit();

    expect(consoleError).toHaveBeenCalledTimes(1);
    expect(loggedLabels()[0]).toMatch(/^\[init\] setupFilterListeners failed:/);
  });

  it("isolates a failure in the pre-load block too", async () => {
    // setupComboboxes runs before the data load, so an unguarded throw there
    // would cost every later step, not just the comboboxes.
    breakStep("setupComboboxes");
    await runInit();

    const survivors = ALL_STEPS.filter((name) => name !== "setupComboboxes");
    expect(stepsThatRan(survivors)).toEqual(survivors);
    expect(loggedLabels()[0]).toMatch(/^\[init\] setupComboboxes failed:/);
  });

  it("survives a localStorage that throws", async () => {
    // Blocked storage (Safari private mode, cookies disabled) makes getItem
    // throw rather than return null. restorePageSize is init()'s first
    // statement, so an unguarded throw there costs every step — the same bug
    // class as above, reached without a missing element.
    //
    // Replacing the whole global rather than spying on Storage.prototype:
    // happy-dom copies each prototype method onto the localStorage instance as
    // a bound own property the first time it is read, so a prototype spy is
    // invisible to every later call — and its restore is swallowed by the
    // instance's Proxy, leaking a throwing getItem into the rest of the file.
    vi.stubGlobal("localStorage", {
      getItem: () => {
        throw new Error("storage disabled");
      },
    });
    try {
      await runInit();

      expect(stepsThatRan()).toEqual(ALL_STEPS);
      expect(loggedLabels()).toEqual([
        expect.stringMatching(/^\[init\] restorePageSize failed:/),
      ]);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("isolates each failure independently, not just the first", async () => {
    breakStep("setupModalListeners");
    breakStep("setupDrawerListeners");
    await runInit();

    const survivors = ALL_STEPS.filter(
      (name) =>
        name !== "setupModalListeners" && name !== "setupDrawerListeners",
    );
    expect(stepsThatRan(survivors)).toEqual(survivors);
    expect(loggedLabels()).toEqual([
      expect.stringMatching(/^\[init\] setupModalListeners failed:/),
      expect.stringMatching(/^\[init\] setupDrawerListeners failed:/),
    ]);
  });
});
