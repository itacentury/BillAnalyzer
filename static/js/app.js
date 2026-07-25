/**
 * Application entry module.
 *
 * Loaded as `<script type="module">`, so it is deferred and runs after the DOM
 * is parsed. It registers the service worker, runs the initial data load and
 * lets every feature module wire up its own DOM event listeners. No functions
 * are exposed on `window` — all interactions are bound via addEventListener.
 */

import {
  applyFilter,
  setupFilterListeners,
  updateFilterBadge,
} from "./filters.js";
import {
  loadInvoices,
  refreshAllData,
  setupPaginationListeners,
} from "./api.js";
import { setupModalListeners } from "./modals.js";
import { setupInvoiceListListeners } from "./render.js";
import { setupBulkListeners } from "./bulk.js";
import { setupStatsListeners } from "./stats.js";
import { setupImportListeners } from "./import.js";
import { setupCategorizeListeners } from "./categorize.js";
import { setupComboboxes } from "./combobox.js";
import { initToastListeners } from "./toast.js";
import { setupKeyboardListeners } from "./keyboard.js";
import { setupDrawerListeners } from "./drawer.js";
import { setupSheetGestures } from "./sheet.js";
import { setupViewportListeners } from "./viewport.js";
import { setupPageSizeListeners } from "./pagesize.js";
import {
  state,
  PAGE_SIZE_OPTIONS,
  ALL_PAGE_SIZE,
  PAGE_SIZE_STORAGE_KEY,
} from "./state.js";

/**
 * Restore the persisted page size into state before the first load, so the
 * initial fetch already uses the user's chosen size. Invalid or missing values
 * keep the default.
 */
function restorePageSize() {
  const stored = localStorage.getItem(PAGE_SIZE_STORAGE_KEY);
  if (stored === null) return;
  if (stored === "all") {
    state.pageSize = ALL_PAGE_SIZE;
    return;
  }
  const parsed = parseInt(stored, 10);
  if (PAGE_SIZE_OPTIONS.includes(parsed)) state.pageSize = parsed;
}

// Register Service Worker for PWA
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker
      .register("/static/sw.js")
      .then((registration) => {
        console.log("[PWA] Service Worker registered:", registration.scope);
      })
      .catch((error) => {
        console.log("[PWA] Service Worker registration failed:", error);
      });
  });
}

function init() {
  restorePageSize();

  // Instantiate the comboboxes before the first data load, so loadStores() and
  // loadCategories() have live instances to feed options into.
  const reloadOnFilterChange = () => {
    updateFilterBadge();
    loadInvoices();
  };
  setupComboboxes({
    "store-filter": reloadOnFilterChange,
    "type-filter": reloadOnFilterChange,
  });

  applyFilter("month");
  refreshAllData();

  setupFilterListeners();
  setupModalListeners();
  setupInvoiceListListeners();
  setupPaginationListeners();
  setupPageSizeListeners();
  setupBulkListeners();
  setupStatsListeners();
  setupImportListeners();
  setupCategorizeListeners();
  initToastListeners();
  setupKeyboardListeners();
  setupDrawerListeners();
  setupSheetGestures();
  setupViewportListeners();
}

// Module scripts run after parsing, so DOMContentLoaded may already have fired.
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}
