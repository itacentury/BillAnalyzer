/**
 * Single-invoice create/update/delete actions.
 */

import { state } from "./state.js";
import { showToast } from "./dom.js";
import { refreshAllData } from "./api.js";
import { closeAddModal, showConfirmModal } from "./modals.js";

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
      refreshAllData();
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
      refreshAllData();
    } else {
      showToast("Failed to delete", "error");
    }
  } catch {
    showToast("Failed to delete", "error");
  }
}
