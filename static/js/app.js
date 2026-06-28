/**
 * Application entry module.
 *
 * Loaded as `<script type="module">`, so it is deferred and runs after the DOM
 * is parsed. It registers the service worker, wires up DOM event listeners,
 * runs the initial data load, and exposes the inline-handler functions on
 * `window` (inline `onclick`/`onchange` attributes resolve against the global
 * scope, which module scope is not part of).
 */

import { els, debounce } from "./dom.js";
import { loadInvoices, loadStores, loadCategories } from "./api.js";
import {
  applyFilter,
  navigateToPrevious,
  navigateToNext,
  resetToCurrent,
  resetAllFilters,
  updateFilterDisplay,
  updateQuickFilterButtons,
} from "./filters.js";
import {
  openAddModal,
  closeAddModal,
  editInvoice,
  openImportModal,
  closeImportModal,
  closeConfirmModal,
  addItemRow,
  removeItemRow,
  calculateTotal,
} from "./modals.js";
import { saveInvoice, deleteInvoice } from "./invoices.js";
import { handleMultipleFiles, removeFile, importJson } from "./import.js";
import { toggleInvoice } from "./render.js";
import {
  toggleInvoiceSelection,
  toggleSelectAll,
  selectAllInvoices,
  deselectAllInvoices,
  openBulkEditModal,
  closeBulkEditModal,
  saveBulkEdit,
  bulkDeleteInvoices,
} from "./bulk.js";
import {
  toggleAdvancedFilters,
  showInvoicesView,
  showStatsView,
} from "./stats.js";
import { state } from "./state.js";

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

function setupEventListeners() {
  const {
    searchInput,
    storeFilter,
    typeFilter,
    dateFrom,
    dateTo,
    sortBy,
    sortOrder,
  } = els();

  searchInput.addEventListener("input", debounce(loadInvoices, 300));
  storeFilter.addEventListener("change", loadInvoices);
  typeFilter.addEventListener("change", loadInvoices);
  // When user manually changes date filters, switch to custom mode
  dateFrom.addEventListener("change", () => {
    if (state.filterMode !== "custom") {
      state.filterMode = "custom";
      updateFilterDisplay();
      updateQuickFilterButtons();
    }
    loadInvoices();
  });
  dateTo.addEventListener("change", () => {
    if (state.filterMode !== "custom") {
      state.filterMode = "custom";
      updateFilterDisplay();
      updateQuickFilterButtons();
    }
    loadInvoices();
  });
  sortBy.addEventListener("change", loadInvoices);
  sortOrder.addEventListener("change", loadInvoices);

  // Dropzone
  const dropzone = document.querySelector('[data-el="dropzone"]');
  const fileInput = document.querySelector('[data-el="file-input"]');

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    const files = e.dataTransfer.files;
    if (files.length > 0) handleMultipleFiles(files);
  });
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) handleMultipleFiles(e.target.files);
  });

  // Calculate total on item input
  document
    .querySelector('[data-el="items-container"]')
    .addEventListener("input", calculateTotal);
}

/**
 * Sync search input values between mobile and desktop search fields.
 */
function syncSearchInputs() {
  const mobileSearch = document.querySelector('[data-el="search"]');
  const desktopSearch = document.querySelector('[data-el="search-desktop"]');

  if (!mobileSearch || !desktopSearch) return;

  mobileSearch.addEventListener("input", () => {
    desktopSearch.value = mobileSearch.value;
  });

  desktopSearch.addEventListener(
    "input",
    debounce(() => {
      mobileSearch.value = desktopSearch.value;
      loadInvoices();
    }, 300),
  );
}

function init() {
  applyFilter("month");
  loadInvoices();
  loadStores();
  loadCategories();
  setupEventListeners();
  syncSearchInputs();
}

// Module scripts run after parsing, so DOMContentLoaded may already have fired.
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}

// Expose the functions referenced by inline HTML handlers to the global scope.
Object.assign(window, {
  toggleInvoice,
  openAddModal,
  closeAddModal,
  editInvoice,
  openImportModal,
  closeImportModal,
  closeConfirmModal,
  addItemRow,
  removeItemRow,
  saveInvoice,
  deleteInvoice,
  removeFile,
  importJson,
  navigateToPrevious,
  navigateToNext,
  resetToCurrent,
  resetAllFilters,
  toggleInvoiceSelection,
  toggleSelectAll,
  selectAllInvoices,
  deselectAllInvoices,
  openBulkEditModal,
  closeBulkEditModal,
  saveBulkEdit,
  bulkDeleteInvoices,
  toggleAdvancedFilters,
  showInvoicesView,
  showStatsView,
});
