/**
 * DOM element cache and generic UI utilities.
 *
 * Leaf module: imports nothing from the app so it can be imported anywhere
 * without creating cycles.
 */

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
 * Format an ISO date string as DD/MM/YYYY.
 */
export function formatDate(dateStr) {
  // Date-only ISO strings parse as UTC midnight; split the parts so the Date
  // is built in local time and the rendered day can't shift by one.
  const [year, month, day] = dateStr.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  return date.toLocaleDateString("en-US", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
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
 * Fill a <datalist> with option suggestions for the given values.
 */
export function populateDatalist(datalist, values) {
  datalist.innerHTML = values
    .map((value) => `<option value="${escapeHtml(value)}">`)
    .join("");
}

/**
 * Show a transient toast notification.
 */
export function showToast(message, type = "success") {
  const container = document.querySelector('[data-el="toast-container"]');
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `
        <span>${type === "success" ? "✓" : "✕"}</span>
        <span>${message}</span>
    `;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

/**
 * Get the current search value from either mobile or desktop input.
 */
export function getSearchValue() {
  const mobileSearch = document.querySelector('[data-el="search"]');
  const desktopSearch = document.querySelector('[data-el="search-desktop"]');

  // Return whichever has a value, prioritizing the visible one based on screen size
  if (window.innerWidth <= 640) {
    return mobileSearch?.value || "";
  }
  return desktopSearch?.value || mobileSearch?.value || "";
}
