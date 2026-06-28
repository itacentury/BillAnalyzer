/**
 * Invoice list rendering and the bulk-action toolbar state.
 *
 * `updateBulkActionToolbar` lives here (rather than in bulk.js) because
 * renderInvoices calls it; keeping it co-located avoids an import cycle —
 * this module depends only on the leaf modules.
 */

import { state, selectedInvoices } from "./state.js";
import { els, escapeHtml, formatDate } from "./dom.js";

export function renderInvoices() {
  const { invoiceList } = els();
  // Clear selection for invoices that are no longer in the list
  const currentIds = new Set(state.invoices.map((inv) => inv.id));
  selectedInvoices.forEach((id) => {
    if (!currentIds.has(id)) {
      selectedInvoices.delete(id);
    }
  });

  if (state.invoices.length === 0) {
    invoiceList.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">📋</div>
                <div class="empty-title">No invoices found</div>
                <div class="empty-text">Adjust your filter criteria or add new invoices.</div>
            </div>
        `;
  } else {
    invoiceList.innerHTML = state.invoices
      .map(
        (invoice) => `
            <div class="invoice-item ${
              selectedInvoices.has(invoice.id) ? "selected" : ""
            }" data-id="${invoice.id}">
                <div class="invoice-header" onclick="toggleInvoice(this)">
                    <label class="invoice-checkbox" onclick="event.stopPropagation()">
                        <input type="checkbox" ${
                          selectedInvoices.has(invoice.id) ? "checked" : ""
                        } onchange="toggleInvoiceSelection(${
                          invoice.id
                        }, this.checked)">
                        <span class="checkbox-mark"></span>
                    </label>
                    <div class="invoice-main">
                        <span class="invoice-date">${formatDate(
                          invoice.date,
                        )}</span>
                        <span class="invoice-store">${escapeHtml(
                          invoice.store,
                        )}</span>
                        ${invoice.category ? `<span class="invoice-type">${escapeHtml(invoice.category)}</span>` : ""}
                    </div>
                    <div class="invoice-meta">
                        <span class="invoice-total">${parseFloat(
                          invoice.total,
                        ).toFixed(2)}</span>
                        <div class="invoice-expand">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <polyline points="6 9 12 15 18 9"/>
                            </svg>
                        </div>
                    </div>
                </div>
                <div class="invoice-details">
                    <div class="items-table">
                        ${invoice.items
                          .map(
                            (item) => `
                            <div class="item-row">
                                <span class="item-name">${escapeHtml(
                                  item.item_name,
                                )}</span>
                                <span class="item-price">€${parseFloat(
                                  item.item_price,
                                ).toFixed(2)}</span>
                            </div>
                        `,
                          )
                          .join("")}
                    </div>
                    <div class="invoice-actions">
                        <button class="btn btn-secondary btn-sm" onclick="editInvoice(${
                          invoice.id
                        })" style="margin-right: 0.5rem;">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                            </svg>
                            Edit
                        </button>
                        <button class="btn btn-danger btn-sm" onclick="deleteInvoice(${
                          invoice.id
                        })">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <polyline points="3 6 5 6 21 6"/>
                                <path d="m19 6-.867 12.142A2 2 0 0 1 16.138 20H7.862a2 2 0 0 1-1.995-1.858L5 6"/>
                                <path d="M10 11v6"/>
                                <path d="M14 11v6"/>
                                <path d="m8 6 .544-1.632A2 2 0 0 1 10.442 3h3.116a2 2 0 0 1 1.898 1.368L16 6"/>
                            </svg>
                            Delete
                        </button>
                    </div>
                </div>
            </div>
        `,
      )
      .join("");
  }

  // Calculate total sum of displayed invoices
  const totalSum = state.invoices.reduce(
    (sum, invoice) => sum + parseFloat(invoice.total),
    0,
  );

  document.querySelector('[data-el="results-count"]').textContent = `${
    state.invoices.length
  } invoice${state.invoices.length !== 1 ? "s" : ""}`;
  document.querySelector('[data-el="results-total"]').textContent =
    totalSum.toFixed(2);

  updateBulkActionToolbar();
}

export function toggleInvoice(element) {
  const item = element.closest(".invoice-item");
  item.classList.toggle("expanded");
}

export function updateBulkActionToolbar() {
  const toolbar = document.querySelector('[data-el="bulk-action-toolbar"]');
  const count = selectedInvoices.size;

  if (count > 0) {
    toolbar.classList.add("visible");
    document.querySelector('[data-el="selected-count"]').textContent = count;
  } else {
    toolbar.classList.remove("visible");
  }

  // Update "select all" checkbox state
  const selectAllCheckbox = document.querySelector(
    '[data-el="select-all-checkbox"] input',
  );
  if (selectAllCheckbox && state.invoices.length > 0) {
    const allSelected = state.invoices.every((inv) =>
      selectedInvoices.has(inv.id),
    );
    const someSelected = state.invoices.some((inv) =>
      selectedInvoices.has(inv.id),
    );

    selectAllCheckbox.checked = allSelected;
    selectAllCheckbox.indeterminate = someSelected && !allSelected;
  } else if (selectAllCheckbox) {
    selectAllCheckbox.checked = false;
    selectAllCheckbox.indeterminate = false;
  }
}
