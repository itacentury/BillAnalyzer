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
  pageSize: 50, // Invoices requested per page
  totalCount: 0, // Total invoices matching the active filters
  totalSum: 0, // Sum of totals across all matching invoices
  currentDate: new Date(), // Current date for navigation reference
  editingInvoiceId: null, // Track if we're editing an invoice
  filterMode: "month", // 'week', 'month', 'year', 'all', 'custom'
  currentView: "invoices", // 'invoices' or 'stats'
  categoryChart: null, // Chart.js instance for category doughnut
  storeChart: null, // Chart.js instance for store bar chart
  pendingFiles: [], // Staged JSON files for import
  confirmModalResolve: null, // Promise resolver for the confirm modal
};

// Track selected invoice IDs for bulk operations (mutated, never reassigned).
export const selectedInvoices = new Set();

// Chart.js color palette matching app theme.
export const chartColors = [
  "#3b82f6", // blue (accent)
  "#22c55e", // green (success)
  "#f59e0b", // amber
  "#ef4444", // red (danger)
  "#8b5cf6", // violet
  "#ec4899", // pink
  "#06b6d4", // cyan
  "#f97316", // orange
  "#84cc16", // lime
  "#6366f1", // indigo
];
