/**
 * Server I/O for the invoice list and the store/category filter dropdowns.
 */

import { state, selectedInvoices } from "./state.js";
import {
  els,
  getSearchValue,
  populateDatalist,
  populateDropdown,
  showToast,
} from "./dom.js";
import { renderInvoices } from "./render.js";
import { loadStats } from "./stats.js";

// Cancels the in-flight invoice request when a newer one supersedes it, so a
// slower earlier response can't render over a newer one (out-of-order results
// on rapid filter/search/sort/pagination changes).
let inFlightController = null;

/**
 * Reload the invoice list plus the store and category lookups in one call.
 */
export function refreshAllData() {
  loadInvoices();
  loadStores();
  loadCategories();
}

/**
 * Load the invoice list, resetting to the first page. Use this for any filter,
 * search, sort or post-mutation refresh; use `goToPage` for pagination.
 */
export async function loadInvoices() {
  // A new query invalidates the selection: clearing here (but not in goToPage /
  // reloadCurrentPage) drops out-of-view ids on filter/search/sort/period
  // changes while preserving cross-page "select all" within a fixed filter set.
  selectedInvoices.clear();
  state.page = 1;
  await fetchInvoices();
}

/**
 * Load a specific invoice-list page without touching the active filters.
 */
export async function goToPage(page) {
  state.page = page;
  await fetchInvoices();
}

/**
 * Reload the current page after a mutation, preserving the user's position.
 * If a deletion emptied the current page, step back to the last page with rows.
 */
export async function reloadCurrentPage() {
  await fetchInvoices();
  const totalPages = Math.max(1, Math.ceil(state.totalCount / state.pageSize));
  if (state.page > totalPages) {
    state.page = totalPages;
    await fetchInvoices();
  }
}

/**
 * Assemble the active filter/search/sort params shared by the list endpoint and
 * the filtered-ids endpoint, so both always see the same query.
 */
function buildFilterParams() {
  const { storeFilter, typeFilter, dateFrom, dateTo, sortBy, sortOrder } =
    els();
  return new URLSearchParams({
    search: getSearchValue(),
    store: storeFilter.value,
    category: typeFilter.value,
    date_from: dateFrom.value,
    date_to: dateTo.value,
    sort_by: sortBy.value,
    sort_order: sortOrder.value,
  });
}

/**
 * Fetch the ids of every invoice matching the active filters (all pages).
 * Backs cross-page "select all".
 */
export async function fetchFilteredIds() {
  const response = await fetch(`/api/invoices/ids?${buildFilterParams()}`);
  const data = await response.json();
  return data.ids;
}

async function fetchInvoices() {
  const params = buildFilterParams();
  params.set("page", state.page);
  params.set("page_size", state.pageSize);

  inFlightController?.abort();
  const controller = new AbortController();
  inFlightController = controller;

  try {
    const response = await fetch(`/api/invoices?${params}`, {
      signal: controller.signal,
    });
    if (!response.ok) throw new Error(`Request failed: ${response.status}`);
    const data = await response.json();
    state.invoices = data.invoices;
    state.page = data.page;
    state.totalCount = data.total_count;
    state.totalSum = data.total_sum;
    renderInvoices();

    // Also refresh stats if in stats view
    if (state.currentView === "stats") {
      loadStats();
    }
  } catch (error) {
    if (error.name === "AbortError") return;
    showToast("Failed to load invoices", "error");
  } finally {
    if (inFlightController === controller) inFlightController = null;
  }
}

/**
 * Wire the pagination control via event delegation on its stable container.
 */
export function setupPaginationListeners() {
  const container = document.querySelector('[data-el="pagination"]');
  if (!container) return;

  container.addEventListener("click", (event) => {
    const button = event.target.closest("[data-action]");
    if (!button || button.disabled) return;
    if (button.dataset.action === "page-prev") goToPage(state.page - 1);
    else if (button.dataset.action === "page-next") goToPage(state.page + 1);
  });
}

export async function loadStores() {
  const { storeFilter } = els();
  try {
    const previousValue = storeFilter.value;
    const response = await fetch("/api/stores");
    const stores = await response.json();

    populateDropdown(storeFilter, stores, "All Stores");

    // Restore filter or jump to next store if previous one no longer exists
    if (previousValue) {
      if (stores.includes(previousValue)) {
        storeFilter.value = previousValue;
      } else if (stores.length > 0) {
        // Find next store alphabetically, or last one if none found
        const nextStore =
          stores.find((s) => s > previousValue) || stores[stores.length - 1];
        storeFilter.value = nextStore;
        loadInvoices();
      }
    }
  } catch (error) {
    console.error("Error loading stores:", error);
  }
}

export async function loadCategories() {
  try {
    const typeFilter = document.querySelector('[data-el="type-filter"]');
    const previousValue = typeFilter ? typeFilter.value : "";

    const response = await fetch("/api/categories");
    const categories = await response.json();

    // Populate type filter dropdown
    if (typeFilter) {
      populateDropdown(typeFilter, categories, "All Categories");

      // Restore filter or reset to "All Categories" if category no longer exists
      if (previousValue) {
        if (categories.includes(previousValue)) {
          typeFilter.value = previousValue;
        } else {
          typeFilter.value = "";
          loadInvoices();
        }
      }
    }

    // Populate datalist suggestions for add/edit form
    const typeSuggestions = document.getElementById("type-suggestions");
    if (typeSuggestions) {
      populateDatalist(typeSuggestions, categories);
    }
  } catch (error) {
    console.error("Error loading categories:", error);
  }
}
