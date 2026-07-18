/**
 * Shared mutable application state.
 *
 * Reassigned values live as properties on the `state` object because ES module
 * import bindings are read-only — a module cannot reassign an imported `let`,
 * but it can mutate a property of an imported object.
 */

export const state = {
  invoices: [],
  page: 1, // Current invoice-list page (1-based)
  pageSize: 25, // Invoices requested per page
  totalCount: 0, // Total invoices matching the active filters
  totalSum: 0, // Sum of totals across all matching invoices
  currentDate: new Date(), // Current date for navigation reference
  editingInvoiceId: null, // Track if we're editing an invoice
  filterMode: "month", // 'week', 'month', 'year', 'all', 'custom'
  currentView: "invoices", // 'invoices' or 'stats'
  categoryChart: null, // Chart.js instance for category doughnut
  storeChart: null, // Chart.js instance for store bar chart
  pendingFiles: [], // Staged JSON files for import
  importErrors: [], // Invalid entries from the last import (index/field/message/value)
};

// Track selected invoice IDs for bulk operations (mutated, never reassigned).
export const selectedInvoices = new Set();

// Selectable invoice-per-page counts offered by the pagination size selector.
export const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

// Sentinel page size for the "All" option: larger than any realistic result set,
// so the server clamps it to MAX_PAGE_SIZE (200) and returns a single page.
export const ALL_PAGE_SIZE = 100000;

// localStorage key persisting the chosen page size across sessions.
export const PAGE_SIZE_STORAGE_KEY = "summa.pageSize";

// Chart.js color palette: warm-sand chart tones (--chart-1…8), donut/bar order.
export const chartColors = [
  "#c9a87c",
  "#a8bfa0",
  "#d9a48a",
  "#b5a184",
  "#c4b3d6",
  "#d6bfa0",
  "#a3c2c2",
  "#e0cdb0",
];
