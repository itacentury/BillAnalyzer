/**
 * Invoice list rendering and the bulk-action toolbar state.
 *
 * `updateBulkActionToolbar` lives here (rather than in bulk.js) because
 * renderInvoices calls it; keeping it co-located avoids an import cycle —
 * this module depends only on the leaf modules.
 */

import { state, selectedInvoices } from "./state.js";
import { els, escapeHtml, formatCurrency, formatDate } from "./dom.js";
import { editInvoice } from "./modals.js";
import { deleteInvoice } from "./invoices.js";
import { toggleInvoiceSelection } from "./bulk.js";

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
                <div class="invoice-header">
                    <label class="invoice-checkbox">
                        <input type="checkbox" ${
                          selectedInvoices.has(invoice.id) ? "checked" : ""
                        }>
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
                        <span class="invoice-total">${formatCurrency(
                          invoice.total,
                        )}</span>
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
                                <span class="item-price">€${formatCurrency(
                                  item.item_price,
                                )}</span>
                            </div>
                        `,
                          )
                          .join("")}
                    </div>
                    <div class="invoice-actions">
                        <button class="btn btn-secondary btn-sm" data-action="edit" style="margin-right: 0.5rem;">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                            </svg>
                            Edit
                        </button>
                        <button class="btn btn-danger btn-sm" data-action="delete">
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

  // Summary reflects the whole filtered set (server totals), not just this page
  document.querySelector('[data-el="results-count"]').textContent = `${
    state.totalCount
  } invoice${state.totalCount !== 1 ? "s" : ""}`;
  document.querySelector('[data-el="results-total"]').textContent =
    formatCurrency(state.totalSum);

  renderPagination();
  updateBulkActionToolbar();
}

/**
 * Render the Prev/Next pagination control. The container is a sibling of the
 * invoice list (which is fully replaced on each render), so it persists.
 */
function renderPagination() {
  const container = document.querySelector('[data-el="pagination"]');
  if (!container) return;

  if (state.totalCount === 0) {
    container.innerHTML = "";
    return;
  }

  const totalPages = Math.max(1, Math.ceil(state.totalCount / state.pageSize));
  container.innerHTML = `
    <button class="btn btn-secondary btn-sm" data-action="page-prev" ${
      state.page <= 1 ? "disabled" : ""
    }>Previous</button>
    <span class="pagination-info">Page ${state.page} of ${totalPages}</span>
    <button class="btn btn-secondary btn-sm" data-action="page-next" ${
      state.page >= totalPages ? "disabled" : ""
    }>Next</button>
  `;
}

export function toggleInvoice(element) {
  const item = element.closest(".invoice-item");
  item.classList.toggle("expanded");
}

/**
 * Wire the invoice list via event delegation so runtime-rendered rows need no
 * per-row listeners. One click and one change listener on the stable container
 * dispatch by the clicked element's `data-action` / its `.invoice-item[data-id]`.
 */
export function setupInvoiceListListeners() {
  const { invoiceList } = els();

  invoiceList.addEventListener("click", (event) => {
    const actionButton = event.target.closest("[data-action]");
    if (actionButton) {
      const id = Number(actionButton.closest(".invoice-item").dataset.id);
      if (actionButton.dataset.action === "edit") editInvoice(id);
      else if (actionButton.dataset.action === "delete") deleteInvoice(id);
      return;
    }

    // Clicking the checkbox must not toggle the row (replaces stopPropagation)
    if (event.target.closest(".invoice-checkbox")) return;

    const header = event.target.closest(".invoice-header");
    if (header) toggleInvoice(header);
  });

  invoiceList.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[type="checkbox"]');
    if (!checkbox) return;
    const id = Number(checkbox.closest(".invoice-item").dataset.id);
    toggleInvoiceSelection(id, checkbox.checked);
  });
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
