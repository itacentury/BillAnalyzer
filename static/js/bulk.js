/**
 * Multi-select, bulk-edit and bulk-delete behavior over the invoice list.
 */

import { state, selectedInvoices } from "./state.js";
import { escapeHtml, showToast } from "./dom.js";
import { loadInvoices, loadStores, loadCategories } from "./api.js";
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

export function toggleSelectAll(isSelected) {
  if (isSelected) {
    state.invoices.forEach((invoice) => selectedInvoices.add(invoice.id));
  } else {
    selectedInvoices.clear();
  }
  renderInvoices();
}

export function selectAllInvoices() {
  state.invoices.forEach((invoice) => selectedInvoices.add(invoice.id));
  renderInvoices();
}

export function deselectAllInvoices() {
  selectedInvoices.clear();
  renderInvoices();
}

export function openBulkEditModal() {
  if (selectedInvoices.size === 0) return;

  // Get the store names and categories of selected invoices
  const selectedStores = new Set();
  const selectedCategories = new Set();
  state.invoices.forEach((invoice) => {
    if (selectedInvoices.has(invoice.id)) {
      selectedStores.add(invoice.store);
      if (invoice.category) {
        selectedCategories.add(invoice.category);
      }
    }
  });

  // Pre-fill with the common store name if all selected have the same store
  const storeInput = document.querySelector('[data-el="bulk-edit-store"]');
  if (selectedStores.size === 1) {
    storeInput.value = [...selectedStores][0];
  } else {
    storeInput.value = "";
    storeInput.placeholder = `${selectedStores.size} different stores`;
  }

  // Pre-fill with the common category if all selected have the same category
  const categoryInput = document.querySelector(
    '[data-el="bulk-edit-category"]',
  );
  if (selectedCategories.size === 1) {
    categoryInput.value = [...selectedCategories][0];
  } else if (selectedCategories.size > 1) {
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
    datalist.innerHTML = categories
      .map((cat) => `<option value="${escapeHtml(cat)}">`)
      .join("");
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
      loadInvoices();
      loadStores();
      loadCategories();
    } else {
      showToast("Failed to update", "error");
    }
  } catch {
    showToast("Failed to update", "error");
  }
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
      loadInvoices();
      loadStores();
    } else {
      showToast("Failed to delete", "error");
    }
  } catch {
    showToast("Failed to delete", "error");
  }
}
