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
import { setupComboboxes } from "./combobox.js";
import { initToastListeners } from "./toast.js";
import { setupKeyboardListeners } from "./keyboard.js";
import { setupDrawerListeners } from "./drawer.js";
import { setupSheetGestures } from "./sheet.js";
import { setupViewportListeners } from "./viewport.js";

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
  setupBulkListeners();
  setupStatsListeners();
  setupImportListeners();
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
