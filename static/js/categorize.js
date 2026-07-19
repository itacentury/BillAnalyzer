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
import { buildFilterParams, refreshAllData } from "./api.js";
import { escapeHtml, formatCurrency } from "./dom.js";
import { showUndoToast, showErrorToast } from "./toast.js";
import { renderInvoices, restoreRows } from "./render.js";
import { lockScroll, unlockScroll } from "./modals.js";
import { createCombobox } from "./combobox.js";

// Per-open review state, reset every time the modal opens.
let controller = null; // aborts the in-flight suggest request on cancel
let rows = []; // suggestion rows from the server
let selected = []; // parallel: whether each row is included
let rowCategory = []; // parallel: the (possibly edited) category per row
let rowIsNew = []; // parallel: whether that category is new
let existingLower = new Set(); // lowercased existing categories, for is-new recompute

/**
 * Show/hide the trigger button and its badge from `state.uncategorizedCount`.
 * Called after every invoice-list load so the badge stays live.
 */
export function updateAiTriggerBadge() {
  const button = document.querySelector('[data-el="ai-categories-trigger"]');
  if (!button) return;
  const count = state.uncategorizedCount || 0;
  button.hidden = count === 0;
  const badge = button.querySelector('[data-el="ai-categories-badge"]');
  if (badge) badge.textContent = count;
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

function groupLabel(group) {
  return group.qty > 1
    ? `${escapeHtml(group.name)} ×${group.qty}`
    : escapeHtml(group.name);
}

/** The item part of the one-line summary (without the trailing amount). */
function summaryItems(groups) {
  if (groups.length === 0) return "No items";
  if (groups.length === 1) return `${escapeHtml(groups[0].name)} · 1 item`;
  const shown = groups.slice(0, 2).map(groupLabel).join(", ");
  const more = groups.length - 2;
  return more > 0
    ? `${shown} <b class="categorize-more">+${more} more</b>`
    : shown;
}

function itemsCardHtml(groups) {
  return groups
    .map(
      (group) => `
        <div class="categorize-item">
          <span class="categorize-item-name">${escapeHtml(group.name)}</span>
          <span class="categorize-item-qty">${group.qty > 1 ? `×${group.qty}` : ""}</span>
          <span class="categorize-item-price">€${formatCurrency(group.price)}</span>
        </div>`,
    )
    .join("");
}

// --- Combobox markup (mirrors the Jinja `combobox` macro, category variant) --

function comboboxMarkup(index) {
  return `
    <div class="combobox" data-combobox data-kind="category" data-empty-label="No Category" data-allow-create="true">
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
    <span class="categorize-new-badge">NEW</span>
  `;
}

// --- Modal state rendering ---------------------------------------------------

function contentEl() {
  return document.querySelector('[data-el="categorize-content"]');
}

function renderLoading() {
  const count = state.uncategorizedCount || 0;
  setFooterVisible(false);
  contentEl().innerHTML = `
    <div class="categorize-banner categorize-banner-loading">
      <span class="categorize-spinner"></span>
      <span>Analyzing ${count} invoice${count !== 1 ? "s" : ""}…</span>
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
  const itemsLabel = `${groups.length} item${groups.length !== 1 ? "s" : ""}`;
  return `
    <div class="categorize-row" data-index="${index}">
      <label class="categorize-check">
        <input type="checkbox" data-el="categorize-row-check" checked aria-label="Include ${escapeHtml(row.store)}" />
        <span class="categorize-check-mark"></span>
      </label>
      <div class="categorize-info">
        <button type="button" class="categorize-rowhead" data-action="toggle-items">
          <span class="categorize-topline">
            <span class="categorize-store">${escapeHtml(row.store)}</span>
            <span class="categorize-amount">€${formatCurrency(row.total)}</span>
          </span>
          <span class="categorize-summary">
            <span class="categorize-summary-items">${summaryItems(groups)}</span>
            <span class="categorize-summary-amount"> · €${formatCurrency(row.total)}</span>
          </span>
        </button>
        <div class="categorize-items">
          <div class="categorize-items-head">
            <span>${itemsLabel} · €${formatCurrency(row.total)}</span>
            <button type="button" class="categorize-showless" data-action="toggle-items">Show less</button>
          </div>
          <div class="categorize-items-card">${itemsCardHtml(groups)}</div>
        </div>
      </div>
      <div class="categorize-combobox">${comboboxMarkup(index)}</div>
    </div>
  `;
}

function renderReview(data, categories) {
  rows = data.suggestions;
  selected = rows.map(() => true);
  rowCategory = rows.map((row) => row.category || "");
  rowIsNew = rows.map((row) => Boolean(row.is_new));
  existingLower = new Set(categories.map((category) => category.toLowerCase()));

  if (rows.length === 0) {
    setFooterVisible(false);
    contentEl().innerHTML = `
      <div class="categorize-banner">Nothing to categorize in this period.</div>`;
    return;
  }

  const cappedNote =
    data.total > data.count
      ? `<div class="categorize-note">First ${data.count} of ${data.total} — run again for the rest.</div>`
      : "";

  contentEl().innerHTML = `
    ${cappedNote}
    <div class="categorize-selectall">
      <label class="categorize-check">
        <input type="checkbox" data-el="categorize-select-all" aria-label="Select all" />
        <span class="categorize-check-mark"></span>
      </label>
      <span class="categorize-selectall-label">Select all</span>
      <button type="button" class="categorize-deselect" data-action="categorize-deselect">Deselect all</button>
    </div>
    <div class="categorize-rows">
      ${rows.map((row, index) => rowHtml(row, index)).join("")}
    </div>
  `;

  // Instantiate each row's category combobox and prefill it with the suggestion.
  contentEl()
    .querySelectorAll(".categorize-combobox")
    .forEach((wrap, index) => {
      const root = wrap.querySelector(".combobox");
      const combobox = createCombobox(root, {
        onChange: (value) => {
          rowCategory[index] = value;
          rowIsNew[index] =
            value !== "" && !existingLower.has(value.toLowerCase());
          updateRowNewBadge(index);
          refreshFooter();
        },
      });
      combobox.setOptions(categories);
      combobox.setValue(rowCategory[index]);
      updateRowNewBadge(index);
    });

  setFooterVisible(true);
  refreshFooter();
}

function updateRowNewBadge(index) {
  const wrap = contentEl().querySelectorAll(".categorize-combobox")[index];
  if (wrap) wrap.classList.toggle("is-new", rowIsNew[index]);
}

// --- Selection, accordion, footer -------------------------------------------

function setRowSelected(index, value) {
  selected[index] = value;
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
  rows.forEach((_, index) => setRowSelected(index, value));
}

/** Accordion: expand one row's item list, collapsing any other open row. */
function toggleItems(index) {
  const rowsEls = contentEl().querySelectorAll(".categorize-row");
  rowsEls.forEach((row) => {
    const isTarget = Number(row.dataset.index) === index;
    row.classList.toggle(
      "expanded",
      isTarget && !row.classList.contains("expanded"),
    );
  });
}

function refreshFooter() {
  const total = rows.length;
  const selectedCount = selected.filter(Boolean).length;

  const newCategories = new Set();
  selected.forEach((isSelected, index) => {
    if (isSelected && rowCategory[index] && rowIsNew[index]) {
      newCategories.add(rowCategory[index].toLowerCase());
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
  rows = [];
  selected = [];
  rowCategory = [];
  rowIsNew = [];
}

async function openCategorize() {
  openModalShell();
  renderLoading();

  controller = new AbortController();
  const current = controller;
  try {
    const [suggestResponse, categoriesResponse] = await Promise.all([
      fetch(`/api/invoices/categorize-suggest?${buildFilterParams()}`, {
        method: "POST",
        signal: current.signal,
      }),
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

    const period =
      document.querySelector('[data-el="month-display"]')?.textContent || "";
    document.querySelector('[data-el="categorize-subtitle"]').textContent =
      `${data.total} uncategorized invoice${data.total !== 1 ? "s" : ""}${period ? ` · ${period}` : ""}`;

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
  selected.forEach((isSelected, index) => {
    if (isSelected && rowCategory[index]) {
      accepted.push({
        id: rows[index].invoice_id,
        category: rowCategory[index],
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
 * deselect-all, cancel) and the footer actions.
 */
export function setupCategorizeListeners() {
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
    if (event.target.closest('[data-action="categorize-deselect"]')) {
      toggleAll(false);
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

  const modal = document.querySelector('[data-el="categorize-modal"]');
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
