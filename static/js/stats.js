/**
 * View switching (invoices/stats), the mobile advanced-filters toggle, and the
 * statistics rendering with Chart.js (the global `Chart` UMD from the CDN).
 */

import { state, chartColors } from "./state.js";
import { els, escapeHtml, formatCurrency, showToast } from "./dom.js";

/**
 * Toggle the visibility of advanced filters on mobile.
 */
export function toggleAdvancedFilters() {
  const collapsible = document.querySelector('[data-el="filters-collapsible"]');
  const toggleBtn = document.querySelector('[data-el="filters-toggle"]');

  collapsible.classList.toggle("visible");
  toggleBtn.classList.toggle("active");
}

/**
 * Switch to the invoices list view.
 */
export function showInvoicesView() {
  state.currentView = "invoices";
  document.querySelector('[data-el="invoices-view"]').style.display = "";
  document.querySelector('[data-el="stats-view"]').style.display = "none";

  // Update toggle buttons
  document.querySelectorAll(".view-toggle-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.view === "invoices");
  });
}

/**
 * Switch to the statistics view and load stats data.
 */
export function showStatsView() {
  state.currentView = "stats";
  document.querySelector('[data-el="invoices-view"]').style.display = "none";
  document.querySelector('[data-el="stats-view"]').style.display = "";

  // Update toggle buttons
  document.querySelectorAll(".view-toggle-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.view === "stats");
  });

  loadStats();
}

/**
 * Wire the invoices/stats view toggle and the mobile advanced-filters toggle.
 */
export function setupStatsListeners() {
  document.querySelector(".view-toggle").addEventListener("click", (event) => {
    const button = event.target.closest(".view-toggle-btn");
    if (!button) return;
    if (button.dataset.view === "stats") showStatsView();
    else showInvoicesView();
  });

  document
    .querySelector('[data-el="filters-toggle"]')
    .addEventListener("click", toggleAdvancedFilters);
}

/**
 * Load statistics data from the API using current date filters.
 */
export async function loadStats() {
  const { dateFrom, dateTo } = els();
  const params = new URLSearchParams({
    date_from: dateFrom.value,
    date_to: dateTo.value,
  });

  try {
    const response = await fetch(`/api/stats?${params}`);
    const data = await response.json();
    renderStats(data);
  } catch {
    showToast("Failed to load statistics", "error");
  }
}

/**
 * Render statistics data including summary cards and charts.
 */
function renderStats(data) {
  const { summary, by_category, by_store, comparison } = data;
  const statsEmpty = document.querySelector('[data-el="stats-empty"]');
  const statsCards = document.querySelector(".stats-cards");
  const statsCharts = document.querySelector(".stats-charts");

  if (summary.total_invoices === 0) {
    statsEmpty.style.display = "";
    statsCards.style.display = "none";
    statsCharts.style.display = "none";
    return;
  }

  statsEmpty.style.display = "none";
  statsCards.style.display = "";
  statsCharts.style.display = "";

  document.querySelector('[data-el="stats-total"]').textContent =
    formatCurrency(summary.total_amount);
  document.querySelector('[data-el="stats-count"]').textContent =
    summary.total_invoices;
  document.querySelector('[data-el="stats-average"]').textContent =
    formatCurrency(summary.average_invoice);

  const changeEl = document.querySelector('[data-el="stats-change"]');
  if (comparison.previous_total > 0) {
    const changePercent = comparison.change_percent;
    const isPositive = changePercent >= 0;
    changeEl.innerHTML = `
      <span class="change-indicator ${isPositive ? "negative" : "positive"}">
        ${isPositive ? "↑" : "↓"} ${Math.abs(changePercent).toFixed(1)}%
      </span>
      <span class="change-label">vs. previous period</span>
    `;
    changeEl.style.display = "";
  } else {
    changeEl.style.display = "none";
  }

  renderCategoryChart(by_category);
  renderStoreChart(by_store);
}

/**
 * Render the category doughnut chart.
 */
function renderCategoryChart(data) {
  const ctx = document.querySelector('[data-el="category-chart"]');
  if (!ctx) return;

  if (state.categoryChart) {
    state.categoryChart.destroy();
  }

  const legendEl = document.querySelector('[data-el="category-legend"]');
  const total = data.reduce((sum, item) => sum + item.amount, 0);
  legendEl.innerHTML = data
    .map((item, i) => {
      const percent = total > 0 ? ((item.amount / total) * 100).toFixed(1) : 0;
      return `
        <div class="legend-item">
          <span class="legend-color" style="background: ${chartColors[i % chartColors.length]}"></span>
          <span class="legend-label">${escapeHtml(item.category)}</span>
          <span class="legend-value">€${formatCurrency(item.amount)}</span>
          <span class="legend-percent">${percent}%</span>
        </div>
      `;
    })
    .join("");

  state.categoryChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: data.map((item) => item.category),
      datasets: [
        {
          data: data.map((item) => item.amount),
          backgroundColor: data.map(
            (_, i) => chartColors[i % chartColors.length],
          ),
          borderWidth: 0,
          hoverOffset: 4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "65%",
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          backgroundColor: "#1a1a1a",
          titleColor: "#fafafa",
          bodyColor: "#a0a0a0",
          borderColor: "#2a2a2a",
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: (context) => {
              const value = context.raw;
              const percent =
                total > 0 ? ((value / total) * 100).toFixed(1) : 0;
              return `€${formatCurrency(value)} (${percent}%)`;
            },
          },
        },
      },
    },
  });
}

/**
 * Render the store horizontal bar chart.
 */
function renderStoreChart(data) {
  const ctx = document.querySelector('[data-el="store-chart"]');
  if (!ctx) return;

  if (state.storeChart) {
    state.storeChart.destroy();
  }

  state.storeChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: data.map((item) => item.store),
      datasets: [
        {
          data: data.map((item) => item.amount),
          backgroundColor: data.map(
            (_, i) => chartColors[i % chartColors.length],
          ),
          borderRadius: 4,
          barThickness: 24,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          backgroundColor: "#1a1a1a",
          titleColor: "#fafafa",
          bodyColor: "#a0a0a0",
          borderColor: "#2a2a2a",
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: (context) => `€${formatCurrency(context.raw)}`,
          },
        },
      },
      scales: {
        x: {
          grid: {
            color: "#2a2a2a",
            drawBorder: false,
          },
          ticks: {
            color: "#666666",
            callback: (value) => `€${value}`,
          },
        },
        y: {
          grid: {
            display: false,
          },
          ticks: {
            color: "#a0a0a0",
          },
        },
      },
    },
  });
}
