/**
 * Modal lifecycle (scroll lock, open/close) for the add/edit, import and
 * confirm dialogs, plus the dynamic item rows of the add/edit form.
 */

import { state } from "./state.js";
import { escapeHtml, showToast } from "./dom.js";
import { saveInvoice } from "./invoices.js";
import { importJson } from "./import.js";

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
  document.querySelector('[data-el="invoice-date"]').valueAsDate = new Date();
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
    showToast("Invoice not found", "error");
    return;
  }

  document.querySelector(
    '[data-el="add-invoice-modal"] .modal-title',
  ).textContent = "Edit Invoice";

  // Fill in the form
  document.querySelector('[data-el="invoice-date"]').value = invoice.date;
  document.querySelector('[data-el="invoice-store"]').value = invoice.store;
  document.querySelector('[data-el="invoice-type"]').value =
    invoice.category || "";

  const itemsContainer = document.querySelector('[data-el="items-container"]');
  itemsContainer.innerHTML = "";

  invoice.items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "item-input-row";
    row.innerHTML = `
            <div class="form-group">
                <label class="form-label">Item Name</label>
                <input type="text" class="form-input item-name" placeholder="Product name" value="${escapeHtml(
                  item.item_name,
                )}">
            </div>
            <div class="form-group">
                <label class="form-label">Price</label>
                <input type="number" step="0.01" class="form-input item-price" placeholder="0.00" value="${
                  item.item_price
                }">
            </div>
            <button type="button" class="btn btn-danger btn-sm" data-action="remove-item" style="margin-bottom: 0.375rem;">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        `;
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
  document.querySelector('[data-el="invoice-type"]').value = "";
  document.querySelector('[data-el="items-container"]').innerHTML = `
        <div class="item-input-row">
            <div class="form-group">
                <label class="form-label">Item Name</label>
                <input type="text" class="form-input item-name" placeholder="Product name">
            </div>
            <div class="form-group">
                <label class="form-label">Price</label>
                <input type="number" step="0.01" class="form-input item-price" placeholder="0.00">
            </div>
            <button type="button" class="btn btn-danger btn-sm" data-action="remove-item" style="margin-bottom: 0.375rem;">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        </div>
    `;
  calculateTotal();
}

export function openImportModal() {
  document.querySelector('[data-el="import-modal"]').classList.add("active");
  lockScroll();
}

/**
 * Show a custom confirmation modal that matches the app design.
 */
export function showConfirmModal(message, title = "Confirm Deletion") {
  document.querySelector('[data-el="confirm-delete-modal-title"]').textContent =
    title;
  document.querySelector(
    '[data-el="confirm-delete-modal-message"]',
  ).textContent = message;
  document
    .querySelector('[data-el="confirm-delete-modal"]')
    .classList.add("active");
  lockScroll();

  return new Promise((resolve) => {
    state.confirmModalResolve = resolve;
  });
}

/**
 * Close the confirmation modal and resolve with the user's choice.
 */
export function closeConfirmModal(confirmed) {
  document
    .querySelector('[data-el="confirm-delete-modal"]')
    .classList.remove("active");
  unlockScroll();
  if (state.confirmModalResolve) {
    state.confirmModalResolve(confirmed);
    state.confirmModalResolve = null;
  }
}

export function closeImportModal() {
  document.querySelector('[data-el="import-modal"]').classList.remove("active");
  unlockScroll();
  document.querySelector('[data-el="json-input"]').value = "";
  document.querySelector('[data-el="file-input"]').value = "";
  state.pendingFiles = [];
  document.querySelector('[data-el="selected-files"]').style.display = "none";
}

export function addItemRow() {
  const container = document.querySelector('[data-el="items-container"]');
  const row = document.createElement("div");
  row.className = "item-input-row";
  row.innerHTML = `
        <div class="form-group">
            <label class="form-label">Item Name</label>
            <input type="text" class="form-input item-name" placeholder="Product name">
        </div>
        <div class="form-group">
            <label class="form-label">Price</label>
            <input type="number" step="0.01" class="form-input item-price" placeholder="0.00">
        </div>
        <button type="button" class="btn btn-danger btn-sm" data-action="remove-item" style="margin-bottom: 0.375rem;">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <line x1="18" y1="6" x2="6" y2="18"/>
                <line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
        </button>
    `;
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
    `€${total.toFixed(2)}`;
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

  // Confirm-delete modal
  const confirmModal = document.querySelector(
    '[data-el="confirm-delete-modal"]',
  );
  confirmModal
    .querySelector(".modal-close")
    .addEventListener("click", () => closeConfirmModal(false));
  confirmModal
    .querySelector('[data-action="cancel"]')
    .addEventListener("click", () => closeConfirmModal(false));
  confirmModal
    .querySelector('[data-action="confirm"]')
    .addEventListener("click", () => closeConfirmModal(true));
}
