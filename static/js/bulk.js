/**
 * Multi-select, bulk-edit and bulk-delete behavior over the invoice list.
 */

import { state, selectedInvoices } from "./state.js";
import { populateDatalist, showToast } from "./dom.js";
import {
  fetchFilteredIds,
  loadCategories,
  loadStores,
  reloadCurrentPage,
} from "./api.js";
import { renderInvoices, updateBulkActionToolbar } from "./render.js";
import { lockScroll, unlockScroll, showConfirmModal } from "./modals.js";

export function toggleInvoiceSelection(invoiceId, isSelected) {
  if (isSelected) {
    selectedInvoices.add(invoiceId);
  } else {
    selectedInvoices.delete(invoiceId);
  }

  // Update visual state of the invoice item
  const invoiceItem = document.querySelector(
    `.invoice-item[data-id="${invoiceId}"]`,
  );
  if (invoiceItem) {
    invoiceItem.classList.toggle("selected", isSelected);
  }

  updateBulkActionToolbar();
}

export async function toggleSelectAll(isSelected) {
  if (isSelected) {
    await selectAllInvoices();
  } else {
    selectedInvoices.clear();
    renderInvoices();
  }
}

/**
 * Select every invoice matching the active filters, across all pages, by
 * fetching the full filtered id set from the server.
 */
export async function selectAllInvoices() {
  try {
    const ids = await fetchFilteredIds();
    ids.forEach((id) => selectedInvoices.add(id));
  } catch {
    showToast("Failed to select all invoices", "error");
    renderInvoices();
    return;
  }
  renderInvoices();
}

export function deselectAllInvoices() {
  selectedInvoices.clear();
  renderInvoices();
}

export function openBulkEditModal() {
  if (selectedInvoices.size === 0) return;

  // Derive the common store/category from the selected invoices to pre-fill the
  // form. Only the current page is loaded client-side, so we can only trust a
  // "common value" when every selected invoice is on this page; otherwise a
  // value shared here might not hold for off-page selections.
  const selectedStores = new Set();
  const selectedCategories = new Set();
  const visibleSelected = state.invoices.filter((invoice) =>
    selectedInvoices.has(invoice.id),
  );
  visibleSelected.forEach((invoice) => {
    selectedStores.add(invoice.store);
    if (invoice.category) {
      selectedCategories.add(invoice.category);
    }
  });
  const allVisible = visibleSelected.length === selectedInvoices.size;

  // Pre-fill with the common store name if all selected are visible and share it
  const storeInput = document.querySelector('[data-el="bulk-edit-store"]');
  if (allVisible && selectedStores.size === 1) {
    storeInput.value = [...selectedStores][0];
  } else if (allVisible) {
    storeInput.value = "";
    storeInput.placeholder = `${selectedStores.size} different stores`;
  } else {
    storeInput.value = "";
    storeInput.placeholder = "Leave empty to keep unchanged";
  }

  // Pre-fill with the common category if all selected are visible and share it
  const categoryInput = document.querySelector(
    '[data-el="bulk-edit-category"]',
  );
  if (allVisible && selectedCategories.size === 1) {
    categoryInput.value = [...selectedCategories][0];
  } else if (allVisible && selectedCategories.size > 1) {
    categoryInput.value = "";
    categoryInput.placeholder = `${selectedCategories.size} different categories`;
  } else {
    categoryInput.value = "";
    categoryInput.placeholder =
      "e.g. Groceries (leave empty to keep unchanged)";
  }

  populateBulkCategorySuggestions();

  document.querySelector('[data-el="bulk-edit-count"]').textContent =
    selectedInvoices.size;
  document.querySelector('[data-el="bulk-edit-modal"]').classList.add("active");
  lockScroll();
  storeInput.focus();
}

async function populateBulkCategorySuggestions() {
  try {
    const response = await fetch("/api/categories");
    const categories = await response.json();
    const datalist = document.getElementById("bulk-category-suggestions");
    populateDatalist(datalist, categories);
  } catch (error) {
    console.error("Error loading categories:", error);
  }
}

export function closeBulkEditModal() {
  document
    .querySelector('[data-el="bulk-edit-modal"]')
    .classList.remove("active");
  unlockScroll();
  document.querySelector('[data-el="bulk-edit-store"]').value = "";
  document.querySelector('[data-el="bulk-edit-category"]').value = "";
}

export async function saveBulkEdit() {
  const newStore = document
    .querySelector('[data-el="bulk-edit-store"]')
    .value.trim();
  const newCategory = document
    .querySelector('[data-el="bulk-edit-category"]')
    .value.trim();

  if (!newStore && !newCategory) {
    showToast("Please fill in at least one field", "error");
    return;
  }

  const ids = [...selectedInvoices];
  const payload = { ids };

  if (newStore) {
    payload.store = newStore;
  }
  if (newCategory) {
    // Only send category if the field has a value
    payload.category = newCategory;
  }

  try {
    const response = await fetch("/api/invoices/bulk-update", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const result = await response.json();
    if (result.success) {
      showToast(`${result.updated} invoice(s) updated`, "success");
      closeBulkEditModal();
      selectedInvoices.clear();
      // Preserve the current page. A bulk edit can rename stores / add or
      // remove categories, so the lookup dropdowns still need refreshing.
      reloadCurrentPage();
      loadStores();
      loadCategories();
    } else {
      showToast("Failed to update", "error");
    }
  } catch {
    showToast("Failed to update", "error");
  }
}

/**
 * Wire the bulk-action toolbar, the bulk-edit modal and the select-all checkbox.
 */
export function setupBulkListeners() {
  const toolbar = document.querySelector('[data-el="bulk-action-toolbar"]');
  toolbar
    .querySelector('[data-action="select-all"]')
    .addEventListener("click", selectAllInvoices);
  toolbar
    .querySelector('[data-action="deselect-all"]')
    .addEventListener("click", deselectAllInvoices);
  toolbar
    .querySelector('[data-action="bulk-edit"]')
    .addEventListener("click", openBulkEditModal);
  toolbar
    .querySelector('[data-action="bulk-delete"]')
    .addEventListener("click", bulkDeleteInvoices);

  const bulkEditModal = document.querySelector('[data-el="bulk-edit-modal"]');
  bulkEditModal
    .querySelector(".modal-close")
    .addEventListener("click", closeBulkEditModal);
  bulkEditModal
    .querySelector('[data-action="cancel"]')
    .addEventListener("click", closeBulkEditModal);
  bulkEditModal
    .querySelector('[data-action="save"]')
    .addEventListener("click", saveBulkEdit);

  const selectAllCheckbox = document.querySelector(
    '[data-el="select-all-checkbox"] input',
  );
  selectAllCheckbox.addEventListener("change", () => {
    toggleSelectAll(selectAllCheckbox.checked);
  });
}

export async function bulkDeleteInvoices() {
  const count = selectedInvoices.size;
  if (count === 0) return;

  const confirmed = await showConfirmModal(
    `Are you sure you want to permanently delete ${count} invoice${
      count !== 1 ? "s" : ""
    }?`,
  );

  if (!confirmed) return;

  const ids = [...selectedInvoices];

  try {
    const response = await fetch("/api/invoices/bulk-delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ids }),
    });

    const result = await response.json();
    if (result.success) {
      showToast(`${result.deleted} invoice(s) deleted`, "success");
      selectedInvoices.clear();
      // Reload the list only; stale lookup options self-heal (see deleteInvoice).
      reloadCurrentPage();
    } else {
      showToast("Failed to delete", "error");
    }
  } catch {
    showToast("Failed to delete", "error");
  }
}
