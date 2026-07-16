/**
 * DOM element cache and generic UI utilities.
 *
 * Leaf module: imports nothing from the app so it can be imported anywhere
 * without creating cycles.
 */

/** Shared mobile breakpoint — keep in sync with the CSS `(width <= 640px)` media queries. */
export const mobileViewport = window.matchMedia("(width <= 640px)");

let cachedEls = null;

/**
 * Return the cached references to the persistent filter/list elements.
 *
 * Queried lazily on first call (not at import time) so module evaluation never
 * depends on the DOM already being parsed.
 */
export function els() {
  if (!cachedEls) {
    cachedEls = {
      invoiceList: document.querySelector('[data-el="invoice-list"]'),
      searchInput: document.querySelector('[data-el="search"]'),
      storeFilter: document.querySelector('[data-el="store-filter"]'),
      typeFilter: document.querySelector('[data-el="type-filter"]'),
      dateFrom: document.querySelector('[data-el="date-from"]'),
      dateTo: document.querySelector('[data-el="date-to"]'),
      sortBy: document.querySelector('[data-el="sort-by"]'),
      sortOrder: document.querySelector('[data-el="sort-order"]'),
      monthDisplay: document.querySelector('[data-el="month-display"]'),
    };
  }
  return cachedEls;
}

/**
 * Debounce a function so it only runs after `wait` ms of inactivity.
 */
export function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

/**
 * Parse an ISO date string into a local-time Date.
 *
 * Date-only ISO strings parse as UTC midnight; split the parts so the Date is
 * built in local time and the rendered day can't shift by one. Anything that
 * isn't a bare YYYY-MM-DD (e.g. an imported datetime) falls back to the
 * permissive Date parser.
 */
function parseLocalDate(dateStr) {
  const parts = dateStr.split("-").map(Number);
  return parts.length === 3 && parts.every(Number.isFinite)
    ? new Date(parts[0], parts[1] - 1, parts[2])
    : new Date(dateStr);
}

/**
 * Format an ISO date string as DD/MM/YYYY.
 */
export function formatDate(dateStr) {
  if (!dateStr) return "";
  return parseLocalDate(dateStr).toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
}

/**
 * Format an ISO date string as DD/MM (the compact mobile list variant).
 */
export function formatDateShort(dateStr) {
  if (!dateStr) return "";
  return parseLocalDate(dateStr).toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
  });
}

/**
 * Escape text for safe insertion into innerHTML.
 */
export function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

const CATEGORY_COLOR_SLUGS = {
  technik: "technik",
  sport: "sport",
  lebensmittel: "lebensmittel",
  bäcker: "baecker",
  bäckerei: "baecker",
  restaurant: "baecker",
  unterkunft: "unterkunft",
};

/**
 * Return an inline style (background + text color) for a category badge.
 *
 * Known categories map to a fixed `--cat-*` pair; anything else is hashed
 * deterministically onto the `--chart-1…8` palette so a given name always keeps
 * the same color. Pure function — only emits CSS-variable references, never the
 * raw category text, so the result is safe to inline into `innerHTML`.
 */
function categoryColorPair(category) {
  const key = category.trim().toLowerCase();
  const mapped = CATEGORY_COLOR_SLUGS[key];
  if (mapped) {
    return { bg: `var(--cat-${mapped})`, text: `var(--cat-${mapped}-text)` };
  }

  let hash = 0;
  for (const char of key) {
    hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  }
  const index = (hash % 8) + 1;
  return { bg: `var(--chart-${index})`, text: "var(--text-primary)" };
}

export function categoryBadgeStyle(category) {
  const { bg, text } = categoryColorPair(category);
  return `background: ${bg}; color: ${text};`;
}

/**
 * Return just the themed background color reference for a category (the color
 * dot used by the category combobox). Emits only a `var(--…)` reference, never
 * the raw category text, so the result is safe to inline.
 */
export function categoryColorVar(category) {
  return categoryColorPair(category).bg;
}

/**
 * Format a numeric amount with exactly two decimals (no currency symbol).
 */
export function formatCurrency(amount) {
  return Number(amount).toFixed(2);
}

/**
 * Fill a <select> with an "all" placeholder followed by the given values.
 */
export function populateDropdown(select, values, allLabel) {
  const options = [`<option value="">${allLabel}</option>`];
  for (const value of values) {
    const safe = escapeHtml(value);
    options.push(`<option value="${safe}">${safe}</option>`);
  }
  select.innerHTML = options.join("");
}

/**
 * Return whether a <select> already contains an option with the given value.
 */
export function hasOption(select, value) {
  return [...select.options].some((option) => option.value === value);
}

/**
 * Get the current search value from the search input.
 */
export function getSearchValue() {
  return document.querySelector('[data-el="search"]')?.value || "";
}
