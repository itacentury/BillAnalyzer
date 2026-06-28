/**
 * Date-range filtering and period navigation (week/month/year/all/custom),
 * including the ISO-week date helpers used only here.
 */

import { state } from "./state.js";
import { els, debounce } from "./dom.js";
import { loadInvoices } from "./api.js";

// Apply a specific filter mode
export function applyFilter(mode) {
  state.filterMode = mode;
  state.currentDate = new Date();
  updateFilterDisplay();
  setDateFiltersForMode();
  updateQuickFilterButtons();
}

// Navigate to previous period based on filter mode
export function navigateToPrevious() {
  if (state.filterMode === "all" || state.filterMode === "custom") return;

  switch (state.filterMode) {
    case "week":
      state.currentDate.setDate(state.currentDate.getDate() - 7);
      break;
    case "month":
      state.currentDate.setMonth(state.currentDate.getMonth() - 1);
      break;
    case "year":
      state.currentDate.setFullYear(state.currentDate.getFullYear() - 1);
      break;
  }
  updateFilterDisplay();
  setDateFiltersForMode();
  loadInvoices();
}

// Navigate to next period based on filter mode
export function navigateToNext() {
  if (state.filterMode === "all" || state.filterMode === "custom") return;

  switch (state.filterMode) {
    case "week":
      state.currentDate.setDate(state.currentDate.getDate() + 7);
      break;
    case "month":
      state.currentDate.setMonth(state.currentDate.getMonth() + 1);
      break;
    case "year":
      state.currentDate.setFullYear(state.currentDate.getFullYear() + 1);
      break;
  }
  updateFilterDisplay();
  setDateFiltersForMode();
  loadInvoices();
}

// Reset to current period for active filter mode
export function resetToCurrent() {
  state.currentDate = new Date();
  updateFilterDisplay();
  setDateFiltersForMode();
  loadInvoices();
}

// Reset all filters back to defaults (current month, no search/store/category)
export function resetAllFilters() {
  const { storeFilter, typeFilter, sortBy, sortOrder } = els();

  // Clear search fields
  const mobileSearch = document.querySelector('[data-el="search"]');
  const desktopSearch = document.querySelector('[data-el="search-desktop"]');
  if (mobileSearch) mobileSearch.value = "";
  if (desktopSearch) desktopSearch.value = "";

  // Reset dropdowns
  storeFilter.value = "";
  typeFilter.value = "";
  sortBy.value = "date";
  sortOrder.value = "desc";

  // Reset to current month
  applyFilter("month");
  loadInvoices();
}

/**
 * Wire the quick-filter buttons, period navigation and the advanced filter
 * inputs, plus the mobile/desktop search field synchronization.
 */
export function setupFilterListeners() {
  // Quick filters: dispatch by the button's data-filter value
  document
    .querySelector(".quick-filters")
    .addEventListener("click", (event) => {
      const button = event.target.closest(".quick-filter-btn");
      if (!button) return;
      applyFilter(button.dataset.filter);
      loadInvoices();
    });

  // Period navigation
  document
    .querySelector('[data-action="nav-prev"]')
    .addEventListener("click", navigateToPrevious);
  document
    .querySelector('[data-action="nav-next"]')
    .addEventListener("click", navigateToNext);
  document
    .querySelector('[data-action="nav-reset"]')
    .addEventListener("click", resetToCurrent);
  document
    .querySelector('[data-action="reset-filters"]')
    .addEventListener("click", resetAllFilters);

  // Advanced filter inputs
  const {
    searchInput,
    storeFilter,
    typeFilter,
    dateFrom,
    dateTo,
    sortBy,
    sortOrder,
  } = els();

  searchInput.addEventListener("input", debounce(loadInvoices, 300));
  storeFilter.addEventListener("change", loadInvoices);
  typeFilter.addEventListener("change", loadInvoices);
  // Manually changing a date filter switches to custom mode
  dateFrom.addEventListener("change", switchToCustomMode);
  dateTo.addEventListener("change", switchToCustomMode);
  sortBy.addEventListener("change", loadInvoices);
  sortOrder.addEventListener("change", loadInvoices);

  syncSearchInputs();
}

// Switch to custom filter mode when a date input is edited directly
function switchToCustomMode() {
  if (state.filterMode !== "custom") {
    state.filterMode = "custom";
    updateFilterDisplay();
    updateQuickFilterButtons();
  }
  loadInvoices();
}

// Keep the mobile and desktop search fields in sync
function syncSearchInputs() {
  const mobileSearch = document.querySelector('[data-el="search"]');
  const desktopSearch = document.querySelector('[data-el="search-desktop"]');

  if (!mobileSearch || !desktopSearch) return;

  mobileSearch.addEventListener("input", () => {
    desktopSearch.value = mobileSearch.value;
  });

  desktopSearch.addEventListener(
    "input",
    debounce(() => {
      mobileSearch.value = desktopSearch.value;
      loadInvoices();
    }, 300),
  );
}

// Update the navigation display based on filter mode
export function updateFilterDisplay() {
  const { monthDisplay } = els();
  const monthNames = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  const navButtons = document.querySelectorAll(".month-nav-btn");
  const resetBtn = document.querySelector(".month-reset-btn");

  switch (state.filterMode) {
    case "week": {
      const weekNum = getISOWeek(state.currentDate);
      const weekYear = getISOWeekYear(state.currentDate);
      monthDisplay.textContent = `W${weekNum} / ${weekYear}`;
      navButtons.forEach((btn) => (btn.style.visibility = "visible"));
      resetBtn.textContent = "Current Week";
      resetBtn.style.display = "";
      break;
    }
    case "month": {
      const monthName = monthNames[state.currentDate.getMonth()];
      const year = state.currentDate.getFullYear();
      monthDisplay.textContent = `${monthName} ${year}`;
      navButtons.forEach((btn) => (btn.style.visibility = "visible"));
      resetBtn.textContent = "Current Month";
      resetBtn.style.display = "";
      break;
    }
    case "year":
      monthDisplay.textContent = `${state.currentDate.getFullYear()}`;
      navButtons.forEach((btn) => (btn.style.visibility = "visible"));
      resetBtn.textContent = "Current Year";
      resetBtn.style.display = "";
      break;
    case "all":
      monthDisplay.textContent = "All Invoices";
      navButtons.forEach((btn) => (btn.style.visibility = "hidden"));
      resetBtn.style.display = "none";
      break;
    case "custom":
      monthDisplay.textContent = "Custom";
      navButtons.forEach((btn) => (btn.style.visibility = "hidden"));
      resetBtn.style.display = "none";
      break;
  }
}

// Update quick filter button active state
export function updateQuickFilterButtons() {
  const buttons = document.querySelectorAll(".quick-filter-btn");
  buttons.forEach((btn) => {
    const btnMode = btn.getAttribute("data-filter");
    btn.classList.toggle("active", btnMode === state.filterMode);
  });
}

// Set date filters based on current mode
function setDateFiltersForMode() {
  const { dateFrom, dateTo } = els();
  switch (state.filterMode) {
    case "week":
      setDateFiltersForWeek(state.currentDate);
      break;
    case "month":
      setDateFiltersForMonth(state.currentDate);
      break;
    case "year":
      setDateFiltersForYear(state.currentDate);
      break;
    case "all":
      dateFrom.value = "";
      dateTo.value = "";
      break;
    case "custom":
      // Don't change the date filters for custom mode
      break;
  }
}

// Calculate ISO week number (weeks start on Monday)
function getISOWeek(date) {
  const d = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()),
  );
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

// Get the year that the ISO week belongs to
function getISOWeekYear(date) {
  const d = new Date(
    Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()),
  );
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  return d.getUTCFullYear();
}

// Get Monday of the week for a given date
function getMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}

// Set date filters for a week (Monday to Sunday)
function setDateFiltersForWeek(date) {
  const { dateFrom, dateTo } = els();
  const monday = getMonday(date);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  dateFrom.value = monday.toLocaleString("sv").split(" ")[0];
  dateTo.value = sunday.toLocaleString("sv").split(" ")[0];
}

// Set date filters for a month
function setDateFiltersForMonth(date) {
  const { dateFrom, dateTo } = els();
  const year = date.getFullYear();
  const month = date.getMonth();

  // First day of the month
  const firstDay = new Date(year, month, 1);
  const firstDayStr = firstDay.toLocaleString("sv").split(" ")[0];

  // Last day of the month
  const lastDay = new Date(year, month + 1, 0);
  const lastDayStr = lastDay.toLocaleString("sv").split(" ")[0];

  dateFrom.value = firstDayStr;
  dateTo.value = lastDayStr;
}

// Set date filters for a year
function setDateFiltersForYear(date) {
  const { dateFrom, dateTo } = els();
  const year = date.getFullYear();

  const firstDay = new Date(year, 0, 1);
  const lastDay = new Date(year, 11, 31);

  dateFrom.value = firstDay.toLocaleString("sv").split(" ")[0];
  dateTo.value = lastDay.toLocaleString("sv").split(" ")[0];
}
