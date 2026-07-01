/**
 * Server I/O for the invoice list and the store/category filter dropdowns.
 */

import { state } from "./state.js";
import {
  els,
  getSearchValue,
  populateDatalist,
  populateDropdown,
  showToast,
} from "./dom.js";
import { renderInvoices } from "./render.js";
import { loadStats } from "./stats.js";

/**
 * Reload the invoice list plus the store and category lookups in one call.
 */
export function refreshAllData() {
  loadInvoices();
  loadStores();
  loadCategories();
}

export async function loadInvoices() {
  const { storeFilter, typeFilter, dateFrom, dateTo, sortBy, sortOrder } =
    els();
  const params = new URLSearchParams({
    search: getSearchValue(),
    store: storeFilter.value,
    category: typeFilter.value,
    date_from: dateFrom.value,
    date_to: dateTo.value,
    sort_by: sortBy.value,
    sort_order: sortOrder.value,
  });

  try {
    const response = await fetch(`/api/invoices?${params}`);
    state.invoices = await response.json();
    renderInvoices();

    // Also refresh stats if in stats view
    if (state.currentView === "stats") {
      loadStats();
    }
  } catch {
    showToast("Failed to load invoices", "error");
  }
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
