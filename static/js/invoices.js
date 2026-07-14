/**
 * Single-invoice create/update/delete actions.
 */

import { state, selectedInvoices } from "./state.js";
import { els, hasOption, showToast } from "./dom.js";
import {
  loadCategories,
  loadInvoices,
  loadStores,
  reloadCurrentPage,
} from "./api.js";
import { closeAddModal, showConfirmModal } from "./modals.js";
import { getCombobox } from "./combobox.js";

export async function saveInvoice() {
  const date = document.querySelector('[data-el="invoice-date"]').value;
  const store = document.querySelector('[data-el="invoice-store"]').value;
  const type =
    document.querySelector('[data-el="invoice-type"]').value.trim() || null;

  if (!date || !store) {
    showToast("Please fill in date and store", "error");
    return;
  }

  const items = [];
  const rows = document.querySelectorAll(".item-input-row");
  rows.forEach((row) => {
    const name = row.querySelector(".item-name").value;
    const price = row.querySelector(".item-price").value;
    if (name && price) {
      items.push({ item_name: name, item_price: price });
    }
  });

  const total = items.reduce(
    (sum, item) => sum + parseFloat(item.item_price),
    0,
  );

  // Capture before closeAddModal() clears state.editingInvoiceId below.
  const isEdit = Boolean(state.editingInvoiceId);

  try {
    let url = "/api/invoices";
    let method = "POST";
    let successMessage = "Invoice saved";

    if (state.editingInvoiceId) {
      url = `/api/invoices/${state.editingInvoiceId}`;
      method = "PUT";
      successMessage = "Invoice updated";
    }

    const response = await fetch(url, {
      method: method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ date, store, category: type, total, items }),
    });

    if (response.ok) {
      showToast(successMessage, "success");
      closeAddModal();

      // Only reload the lookups when this save introduced a new store/category.
      // The category filter is a combobox (hidden input, no .options), so its
      // membership check goes through the combobox API rather than hasOption().
      const { storeFilter } = els();
      const typeCombobox = getCombobox("type-filter");
      if (!hasOption(storeFilter, store)) loadStores();
      if (type && typeCombobox && !typeCombobox.hasOption(type))
        loadCategories();
      // Editing keeps the user on the current page; a new invoice jumps to
      // page 1 so it is visible at the top of the date-descending sort.
      if (isEdit) reloadCurrentPage();
      else loadInvoices();
    } else {
      showToast("Failed to save", "error");
    }
  } catch {
    showToast("Failed to save", "error");
  }
}

export async function deleteInvoice(id) {
  const confirmed = await showConfirmModal(
    "Are you sure you want to delete this invoice?",
  );
  if (!confirmed) return;

  try {
    const response = await fetch(`/api/invoices/${id}`, { method: "DELETE" });
    if (response.ok) {
      showToast("Invoice deleted", "success");
      // Drop the id from any active selection; without render-time pruning it
      // would otherwise linger and inflate the bulk-action count.
      selectedInvoices.delete(id);
      // A store/category option lingering after its last invoice is deleted is
      // cosmetic and self-heals on the next lookup load, so reload the list only.
      reloadCurrentPage();
    } else {
      showToast("Failed to delete", "error");
    }
  } catch {
    showToast("Failed to delete", "error");
  }
}
