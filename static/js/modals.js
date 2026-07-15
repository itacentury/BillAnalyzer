/**
 * Modal lifecycle (scroll lock, open/close) for the add/edit and import dialogs,
 * plus the dynamic item rows of the add/edit form.
 */

import { state } from "./state.js";
import { escapeHtml, formatCurrency } from "./dom.js";
import { showErrorToast } from "./toast.js";
import { saveInvoice } from "./invoices.js";
import { fetchInvoiceItems } from "./api.js";
import { importJson, setImportCorrectionMode } from "./import.js";
import { getCombobox } from "./combobox.js";

/**
 * Build the inner markup for one add/edit item row (name, price, remove button).
 * Pre-fills the inputs when an existing item is passed.
 */
function itemRowInnerHtml(item = null) {
  const nameValue = item ? ` value="${escapeHtml(item.item_name)}"` : "";
  const priceValue = item ? ` value="${item.item_price}"` : "";
  return `
    <div class="form-group">
      <label class="form-label">Item Name</label>
      <input type="text" class="form-input item-name" placeholder="Product name"${nameValue}>
    </div>
    <div class="form-group">
      <label class="form-label">Price</label>
      <input type="number" step="0.01" class="form-input item-price" placeholder="0.00"${priceValue}>
    </div>
    <button type="button" class="btn btn-danger btn-sm" data-action="remove-item">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <line x1="18" y1="6" x2="6" y2="18"/>
        <line x1="6" y1="6" x2="18" y2="18"/>
      </svg>
    </button>
  `;
}

/**
 * Lock body scroll to prevent background scrolling while a modal is open.
 */
export function lockScroll() {
  document.body.style.overflow = "hidden";
}

/**
 * Restore body scroll when no modals are open.
 */
export function unlockScroll() {
  const anyOpen = document.querySelector(".modal-overlay.active");
  if (!anyOpen) {
    document.body.style.overflow = "";
  }
}

export function openAddModal() {
  state.editingInvoiceId = null;
  document.querySelector(
    '[data-el="add-invoice-modal"] .modal-title',
  ).textContent = "New Invoice";
  document
    .querySelector('[data-el="add-invoice-modal"]')
    .classList.add("active");
  lockScroll();
  resetAddForm();
  document.querySelector('[data-el="invoice-date"]').value = new Date()
    .toLocaleString("sv")
    .split(" ")[0];
}

export function closeAddModal() {
  document
    .querySelector('[data-el="add-invoice-modal"]')
    .classList.remove("active");
  unlockScroll();
  state.editingInvoiceId = null;
  resetAddForm();
}

export async function editInvoice(id) {
  state.editingInvoiceId = id;
  const invoice = state.invoices.find((inv) => inv.id === id);

  if (!invoice) {
    showErrorToast("Invoice not found");
    return;
  }

  document.querySelector(
    '[data-el="add-invoice-modal"] .modal-title',
  ).textContent = "Edit Invoice";

  // Fill in the form
  document.querySelector('[data-el="invoice-date"]').value = invoice.date;
  document.querySelector('[data-el="invoice-store"]').value = invoice.store;
  getCombobox("invoice-type").setValue(invoice.category || "");

  // The compact list no longer carries items, so load them on demand.
  let items;
  try {
    items = await fetchInvoiceItems(id);
  } catch {
    // A newer editInvoice() superseded this one; that call owns the error surface.
    if (state.editingInvoiceId !== id) return;
    showErrorToast("Failed to load invoice");
    return;
  }

  // A newer editInvoice() superseded this one while items were loading; its
  // header and editingInvoiceId now own the modal, so don't inject stale items.
  if (state.editingInvoiceId !== id) return;

  const itemsContainer = document.querySelector('[data-el="items-container"]');
  itemsContainer.innerHTML = "";

  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "item-input-row";
    row.innerHTML = itemRowInnerHtml(item);
    itemsContainer.appendChild(row);
  });

  calculateTotal();
  document
    .querySelector('[data-el="add-invoice-modal"]')
    .classList.add("active");
  lockScroll();
}

function resetAddForm() {
  document.querySelector('[data-el="add-form"]').reset();
  getCombobox("invoice-type").setValue("");
  document.querySelector('[data-el="items-container"]').innerHTML =
    `<div class="item-input-row">${itemRowInnerHtml()}</div>`;
  calculateTotal();
}

export function openImportModal() {
  document.querySelector('[data-el="import-modal"]').classList.add("active");
  lockScroll();
}

export function closeImportModal() {
  document.querySelector('[data-el="import-modal"]').classList.remove("active");
  unlockScroll();
  document.querySelector('[data-el="json-input"]').value = "";
  document.querySelector('[data-el="file-input"]').value = "";
  state.pendingFiles = [];
  document.querySelector('[data-el="selected-files"]').style.display = "none";

  // Clear any staged import errors so reopening never shows stale cards.
  const importErrors = document.querySelector('[data-el="import-errors"]');
  importErrors.innerHTML = "";
  importErrors.style.display = "none";
  state.importErrors = [];

  // Restore the fresh-input controls hidden while in correction mode.
  setImportCorrectionMode(false);
}

export function addItemRow() {
  const container = document.querySelector('[data-el="items-container"]');
  const row = document.createElement("div");
  row.className = "item-input-row";
  row.innerHTML = itemRowInnerHtml();
  container.appendChild(row);
}

export function removeItemRow(button) {
  const rows = document.querySelectorAll(".item-input-row");
  if (rows.length > 1) {
    button.closest(".item-input-row").remove();
    calculateTotal();
  }
}

export function calculateTotal() {
  const prices = document.querySelectorAll(".item-price");
  let total = 0;
  prices.forEach((input) => {
    total += parseFloat(input.value) || 0;
  });
  document.querySelector('[data-el="calculated-total"]').textContent =
    `€${formatCurrency(total)}`;
}

/**
 * Wire the modal open/close buttons and the add/edit form, including delegation
 * on the items container for the dynamically added item rows.
 */
export function setupModalListeners() {
  // Header buttons that open the modals
  document
    .querySelector('[data-action="open-add"]')
    .addEventListener("click", openAddModal);
  document
    .querySelector('[data-action="open-import"]')
    .addEventListener("click", openImportModal);

  // Add/edit invoice modal
  const addModal = document.querySelector('[data-el="add-invoice-modal"]');
  addModal
    .querySelector(".modal-close")
    .addEventListener("click", closeAddModal);
  addModal
    .querySelector('[data-action="cancel"]')
    .addEventListener("click", closeAddModal);
  addModal
    .querySelector('[data-action="add-item"]')
    .addEventListener("click", addItemRow);
  addModal
    .querySelector('[data-action="save"]')
    .addEventListener("click", saveInvoice);

  // Item rows are rebuilt at runtime, so delegate from the stable container
  const itemsContainer = document.querySelector('[data-el="items-container"]');
  itemsContainer.addEventListener("input", calculateTotal);
  itemsContainer.addEventListener("click", (event) => {
    const removeButton = event.target.closest('[data-action="remove-item"]');
    if (removeButton) removeItemRow(removeButton);
  });

  // Import modal
  const importModal = document.querySelector('[data-el="import-modal"]');
  importModal
    .querySelector(".modal-close")
    .addEventListener("click", closeImportModal);
  importModal
    .querySelector('[data-action="cancel"]')
    .addEventListener("click", closeImportModal);
  importModal
    .querySelector('[data-action="import"]')
    .addEventListener("click", importJson);
}
