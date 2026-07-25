/**
 * AI category suggestions: the "AI Categories" trigger, the review modal / bottom
 * sheet, and applying the confirmed categories.
 *
 * The endpoint is read-only — it only returns suggestions. Writing goes through
 * the existing `PUT /api/invoices/bulk-update` path (one call per unique target
 * category), reusing the optimistic-update + undo-toast flow from bulk.js so the
 * feature behaves exactly like a bulk edit.
 */

import { state } from "./state.js";
import { refreshAllData } from "./api.js";
import { escapeHtml, formatCurrency } from "./dom.js";
import { showUndoToast, showErrorToast, flushPendingToast } from "./toast.js";
import { renderInvoices, restoreRows } from "./render.js";
import { lockScroll, unlockScroll } from "./modals.js";
import { createCombobox } from "./combobox.js";
import { getAiModel, setupModelPicker } from "./ai-model.js";

// Per-open review state, reset every time the modal opens.
let controller = null; // aborts the in-flight suggest request on cancel
// One object per suggestion row, index-aligned to the rendered rows:
// { suggestion, selected, expanded, category, isNew }.
let reviewRows = [];
let existingLower = new Set(); // lowercased existing categories, for is-new recompute

/**
 * Enable/disable the trigger button and update its badge from the uncategorized
 * invoices on the current page (`state.invoices`), matching what the AI action
 * analyzes. Called after every invoice-list load so it stays live. When nothing
 * is uncategorized the button is greyed out (disabled), not hidden, so it keeps
 * its place in the toolbar.
 */
export function updateAiTriggerBadge() {
  const button = document.querySelector('[data-el="ai-categories-trigger"]');
  if (!button) return;
  const count = state.invoices.filter((invoice) => !invoice.category).length;
  button.disabled = count === 0;
  const badge = button.querySelector('[data-el="ai-categories-badge"]');
  if (badge) badge.textContent = String(count);
}

// --- Item aggregation & summary ---------------------------------------------

/**
 * Collapse repeated line items (the schema has no quantity column) into groups
 * of {name, qty, price}, preserving first-seen order. `qty` is the row count and
 * `price` the summed line prices for that name.
 */
function aggregateItems(items) {
  const groups = new Map();
  items.forEach((item) => {
    const name = item.item_name || "";
    const group = groups.get(name) || { name, qty: 0, price: 0 };
    group.qty += 1;
    group.price += Number(item.item_price) || 0;
    groups.set(name, group);
  });
  return [...groups.values()];
}

function itemsLineHtml(groups) {
  return groups
    .map(
      (group) => `
        <div class="categorize-item-line">
          <span class="categorize-item-line-name">${group.qty > 1 ? `${group.qty}× ` : ""}${escapeHtml(group.name)}</span>
          <span class="categorize-item-line-price">€${formatCurrency(group.price)}</span>
        </div>`,
    )
    .join("");
}

// --- Combobox markup (mirrors the Jinja `combobox` macro, category variant) --

function comboboxMarkup(index) {
  return `
    <span class="categorize-new-badge">NEW</span>
    <div class="combobox" data-combobox data-kind="category" data-empty-label="No Category" data-allow-create="true" data-menu-float="true">
      <input type="hidden" data-el="categorize-cat-${index}" />
      <div class="combobox-control">
        <svg class="combobox-search" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" />
        </svg>
        <span class="combobox-dot" hidden></span>
        <input type="text" class="combobox-input" placeholder="Select or type a category…" role="combobox" aria-expanded="false" autocomplete="off" />
        <svg class="combobox-chevron" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polyline points="6 9 12 15 18 9" />
        </svg>
      </div>
      <ul class="combobox-menu" role="listbox"></ul>
    </div>
  `;
}

// --- Modal state rendering ---------------------------------------------------

function contentEl() {
  return document.querySelector('[data-el="categorize-content"]');
}

function renderLoading() {
  // No count: with caching only new/edited invoices reach the model, and the
  // true analyzed amount is not known until the server responds.
  setFooterVisible(false);
  contentEl().innerHTML = `
    <div class="categorize-banner categorize-banner-loading">
      <span class="categorize-spinner"></span>
      <span>Analyzing invoices…</span>
      <button type="button" class="categorize-banner-action" data-action="categorize-cancel">Cancel</button>
    </div>
  `;
}

function renderNotConfigured() {
  setFooterVisible(false);
  contentEl().innerHTML = `
    <div class="categorize-banner categorize-banner-error">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="9" /><line x1="12" y1="8" x2="12" y2="13" /><line x1="12" y1="16.5" x2="12" y2="16.5" />
      </svg>
      <span><strong>AI categorization not configured.</strong> Set <code>ANTHROPIC_API_KEY</code> in the server environment.</span>
    </div>
  `;
}

function renderError(message) {
  setFooterVisible(false);
  contentEl().innerHTML = `
    <div class="categorize-banner categorize-banner-error">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="12" cy="12" r="9" /><line x1="12" y1="8" x2="12" y2="13" /><line x1="12" y1="16.5" x2="12" y2="16.5" />
      </svg>
      <span>${escapeHtml(message)}</span>
    </div>
  `;
}

function rowHtml(row, index) {
  const groups = aggregateItems(row.items);
  const itemCount = groups.length;
  const hasExpandableItems = itemCount > 1;
  const isExpanded = reviewRows[index].expanded;
  const total = `€${formatCurrency(row.total)}`;
  const metaLead =
    itemCount === 1
      ? `${escapeHtml(groups[0]?.name || "No items")} · 1 item · ${total}`
      : `${itemCount} items · ${total}`;
  const toggleLabel = isExpanded ? "Show less" : "Show items";

  return `
    <div class="categorize-row${isExpanded ? " expanded" : ""}" data-index="${index}">
      <label class="categorize-check">
        <input type="checkbox" data-el="categorize-row-check" checked aria-label="Include ${escapeHtml(row.store)}" />
        <span class="categorize-check-mark"></span>
      </label>
      <div class="categorize-info">
        <span class="categorize-topline">
          <span class="categorize-store">${escapeHtml(row.store)}</span>
          <span class="categorize-amount">${total}</span>
        </span>
        ${
          hasExpandableItems
            ? `<button
                 type="button"
                 class="categorize-meta-button"
                 data-action="toggle-items"
                 data-el="categorize-meta-toggle"
                 aria-expanded="${isExpanded ? "true" : "false"}"
               >
                 <span class="categorize-meta-text">${metaLead}</span>
                 <span class="categorize-meta-sep" aria-hidden="true">·</span>
                 <span class="categorize-meta-toggle" data-el="categorize-toggle-label">${toggleLabel}</span>
               </button>`
            : `<span class="categorize-meta-text">${metaLead}</span>`
        }
        ${
          hasExpandableItems
            ? `<div class="categorize-items">${itemsLineHtml(groups)}</div>`
            : ""
        }
      </div>
      <div class="categorize-combobox">${comboboxMarkup(index)}</div>
    </div>
  `;
}

function renderReview(data, categories) {
  reviewRows = data.suggestions.map((suggestion) => ({
    suggestion,
    selected: true,
    expanded: false,
    category: suggestion.category || "",
    isNew: Boolean(suggestion.is_new),
  }));
  existingLower = new Set(categories.map((category) => category.toLowerCase()));

  if (reviewRows.length === 0) {
    setFooterVisible(false);
    contentEl().innerHTML = `
      <div class="categorize-banner">Nothing to categorize on this page.</div>`;
    return;
  }

  const cappedNote =
    data.total > data.count
      ? `<div class="categorize-note">First ${data.count} of ${data.total} — apply these, then run again for the rest.</div>`
      : "";

  contentEl().innerHTML = `
    ${cappedNote}
    <div class="categorize-selectall">
      <label class="categorize-check">
        <input type="checkbox" data-el="categorize-select-all" aria-label="Select all" />
        <span class="categorize-check-mark"></span>
      </label>
      <span class="categorize-selectall-label">Select all</span>
    </div>
    <div class="categorize-rows">
      ${reviewRows.map((entry, index) => rowHtml(entry.suggestion, index)).join("")}
    </div>
  `;

  // Instantiate each row's category combobox and prefill it with the suggestion.
  contentEl()
    .querySelectorAll(".categorize-combobox")
    .forEach((wrap, index) => {
      const root = wrap.querySelector(".combobox");
      const combobox = createCombobox(root, {
        onChange: (value) => {
          reviewRows[index].category = value;
          reviewRows[index].isNew =
            value !== "" && !existingLower.has(value.toLowerCase());
          updateRowNewBadge(index);
          refreshFooter();
        },
      });
      combobox.setOptions(categories);
      combobox.setValue(reviewRows[index].category);
      updateRowNewBadge(index);
    });

  setFooterVisible(true);
  refreshFooter();
}

function updateRowNewBadge(index) {
  const wrap = contentEl().querySelectorAll(".categorize-combobox")[index];
  if (wrap) wrap.classList.toggle("is-new", reviewRows[index].isNew);
}

// --- Selection, accordion, footer -------------------------------------------

function setRowSelected(index, value) {
  reviewRows[index].selected = value;
  const row = contentEl().querySelector(
    `.categorize-row[data-index="${index}"]`,
  );
  if (row) {
    row.classList.toggle("is-deselected", !value);
    const checkbox = row.querySelector('[data-el="categorize-row-check"]');
    if (checkbox) checkbox.checked = value;
  }
  refreshFooter();
}

function toggleAll(value) {
  reviewRows.forEach((_, index) => setRowSelected(index, value));
}

function toggleItems(index) {
  if (index < 0 || index >= reviewRows.length) return;
  const isExpanded = !reviewRows[index].expanded;
  reviewRows[index].expanded = isExpanded;

  const row = contentEl().querySelector(
    `.categorize-row[data-index="${index}"]`,
  );
  if (!row) return;

  row.classList.toggle("expanded", isExpanded);

  const toggleButton = row.querySelector('[data-el="categorize-meta-toggle"]');
  if (toggleButton) {
    toggleButton.setAttribute("aria-expanded", isExpanded ? "true" : "false");
  }

  const toggleLabel = row.querySelector('[data-el="categorize-toggle-label"]');
  if (toggleLabel) {
    toggleLabel.textContent = isExpanded ? "Show less" : "Show items";
  }
}

function refreshFooter() {
  const total = reviewRows.length;
  const selectedCount = reviewRows.filter((entry) => entry.selected).length;

  const newCategories = new Set();
  reviewRows.forEach((entry) => {
    if (entry.selected && entry.category && entry.isNew) {
      newCategories.add(entry.category.toLowerCase());
    }
  });
  const newCount = newCategories.size;

  let summary = `${selectedCount} of ${total} selected`;
  if (newCount > 0) {
    summary += ` · ${newCount} new categor${newCount === 1 ? "y" : "ies"}`;
  }
  document.querySelector('[data-el="categorize-summary-text"]').textContent =
    summary;

  document.querySelector('[data-el="categorize-apply-count"]').textContent =
    selectedCount;
  document.querySelector('[data-action="categorize-apply"]').disabled =
    selectedCount === 0;

  const selectAll = document.querySelector('[data-el="categorize-select-all"]');
  if (selectAll) {
    selectAll.checked = total > 0 && selectedCount === total;
    selectAll.indeterminate = selectedCount > 0 && selectedCount < total;
  }
}

function setFooterVisible(visible) {
  document.querySelector('[data-el="categorize-footer"]').hidden = !visible;
}

// --- Open / close / fetch ----------------------------------------------------

function openModalShell() {
  const modal = document.querySelector('[data-el="categorize-modal"]');
  document.querySelector('[data-el="categorize-subtitle"]').textContent = "";
  modal.classList.add("active");
  lockScroll();
}

export function closeCategorizeModal() {
  controller?.abort();
  controller = null;
  document
    .querySelector('[data-el="categorize-modal"]')
    .classList.remove("active");
  unlockScroll();
  contentEl().innerHTML = "";
  setFooterVisible(false);
  reviewRows = [];
}

async function openCategorize() {
  openModalShell();
  renderLoading();

  // Finalize any deferred invoice edit/delete first so its new content reaches
  // the server before we ask for suggestions — otherwise a just-edited invoice
  // would still match its cached fingerprint and return the pre-edit suggestion.
  await flushPendingToast();

  await runAnalysis();
}

/**
 * Re-run the model switch inside the open modal: abort the previous request,
 * show the loading banner (which hides the footer/Apply), and fetch fresh
 * suggestions for the currently selected model. No-op when the modal is closed.
 */
async function rerunForModel() {
  const modal = document.querySelector('[data-el="categorize-modal"]');
  if (!modal.classList.contains("active")) return;
  renderLoading();
  await runAnalysis();
}

/**
 * Fetch suggestions for the current filter + model and render the review. The
 * caller shows the loading banner first; this owns the request lifecycle and
 * aborts any in-flight request so a rapid model switch never double-renders.
 */
export async function runAnalysis() {
  controller?.abort();
  controller = new AbortController();
  const current = controller;
  // Scope the analysis to exactly the uncategorized invoices on the current page.
  const ids = state.invoices
    .filter((invoice) => !invoice.category)
    .map((invoice) => invoice.id);
  try {
    const [suggestResponse, categoriesResponse] = await Promise.all([
      fetch(
        `/api/invoices/categorize-suggest?model=${encodeURIComponent(getAiModel())}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ ids }),
          signal: current.signal,
        },
      ),
      fetch("/api/categories"),
    ]);

    if (suggestResponse.status === 503) {
      renderNotConfigured();
      return;
    }
    if (!suggestResponse.ok) {
      const body = await suggestResponse.json().catch(() => ({}));
      renderError(body.error || "Failed to analyze invoices");
      return;
    }

    const data = await suggestResponse.json();
    const categories = await categoriesResponse.json();

    document.querySelector('[data-el="categorize-subtitle"]').textContent =
      `${data.total} uncategorized invoice${data.total !== 1 ? "s" : ""} on this page`;

    renderReview(data, categories);
  } catch (error) {
    if (error.name === "AbortError") return; // user cancelled; modal already closed
    renderError("Failed to analyze invoices");
  } finally {
    if (controller === current) controller = null;
  }
}

// --- Apply -------------------------------------------------------------------

function applyCategories() {
  // Accepted = selected rows that carry a non-empty category.
  const accepted = [];
  reviewRows.forEach((entry) => {
    if (entry.selected && entry.category) {
      accepted.push({
        id: entry.suggestion.invoice_id,
        category: entry.category,
      });
    }
  });
  if (accepted.length === 0) return;

  // Group ids by unique target category → one bulk-update call per category.
  const idsByCategory = new Map();
  accepted.forEach(({ id, category }) => {
    const list = idsByCategory.get(category) || [];
    list.push(id);
    idsByCategory.set(category, list);
  });

  const idToCategory = new Map(
    accepted.map(({ id, category }) => [id, category]),
  );
  const idSet = new Set(idToCategory.keys());
  const count = accepted.length;

  // Optimistically categorize any of these rows visible on the current page;
  // snapshot the old versions so undo can restore them (off-page rows reconcile
  // via refreshAllData on commit).
  const previous = state.invoices.filter((invoice) => idSet.has(invoice.id));
  state.invoices = state.invoices.map((invoice) =>
    idSet.has(invoice.id)
      ? { ...invoice, category: idToCategory.get(invoice.id) }
      : invoice,
  );

  closeCategorizeModal();
  renderInvoices();

  const revert = () => restoreRows(previous);

  const commit = async () => {
    try {
      const responses = await Promise.all(
        [...idsByCategory.entries()].map(([category, ids]) =>
          fetch("/api/invoices/bulk-update", {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ ids, category }),
            keepalive: true,
          }),
        ),
      );
      const bodies = await Promise.all(
        responses.map((response) => response.json().catch(() => ({}))),
      );
      if (!bodies.every((body) => body.success)) {
        showErrorToast("Failed to apply categories");
        revert();
        return;
      }
      // New categories may have been created and existing rows recategorized, so
      // refresh the list plus the store/category lookups.
      refreshAllData();
    } catch {
      showErrorToast("Failed to apply categories");
      revert();
    }
  };

  showUndoToast(`${count} invoice${count !== 1 ? "s" : ""} categorized`, {
    onUndo: revert,
    onCommit: commit,
  });
}

// --- Wiring ------------------------------------------------------------------

/**
 * Wire the trigger button, the review-content delegation (checkboxes, accordion,
 * cancel) and the footer actions.
 */
export function setupCategorizeListeners() {
  const modal = document.querySelector('[data-el="categorize-modal"]');
  if (!modal) return; // AI suggestions disabled: modal partial not rendered

  // Switching the model re-runs the analysis live when the modal is open.
  setupModelPicker(rerunForModel);

  document
    .querySelectorAll('[data-action="open-categorize"]')
    .forEach((button) => button.addEventListener("click", openCategorize));

  const content = contentEl();

  content.addEventListener("click", (event) => {
    // Let the combobox handle its own clicks.
    if (event.target.closest(".categorize-combobox")) return;

    if (event.target.closest('[data-action="categorize-cancel"]')) {
      closeCategorizeModal();
      return;
    }
    const toggle = event.target.closest('[data-action="toggle-items"]');
    if (toggle) {
      const row = toggle.closest(".categorize-row");
      toggleItems(Number(row.dataset.index));
    }
  });

  content.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[type="checkbox"]');
    if (!checkbox) return;
    if (checkbox.dataset.el === "categorize-select-all") {
      toggleAll(checkbox.checked);
      return;
    }
    const row = checkbox.closest(".categorize-row");
    if (row) setRowSelected(Number(row.dataset.index), checkbox.checked);
  });

  modal
    .querySelector(".modal-close")
    .addEventListener("click", closeCategorizeModal);
  modal
    .querySelector('[data-action="categorize-cancel-footer"]')
    .addEventListener("click", closeCategorizeModal);
  modal
    .querySelector('[data-action="categorize-apply"]')
    .addEventListener("click", applyCategories);
}
