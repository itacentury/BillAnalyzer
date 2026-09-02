import { beforeEach, describe, expect, it } from "vitest";

import { toggleAdvancedFilters } from "../../static/js/stats.js";

// The collapsed panel is hidden by CSS only (opacity + a zero-height grid row),
// so `inert` is the sole thing keeping its controls out of the tab order and
// out of hit testing. These tests pin that contract.
const markup = `
  <button class="filter-button" data-el="filters-toggle" aria-expanded="false" aria-controls="filters-panel">Filter</button>
  <button class="filter-sheet-scrim" data-el="filter-sheet-scrim" tabindex="-1"></button>
  <div class="filters-collapsible" id="filters-panel" data-el="filters-collapsible" inert>
    <input type="text" class="combobox-input" data-el="store-filter-input" />
  </div>
`;

const panel = () => document.querySelector('[data-el="filters-collapsible"]');
const toggle = () => document.querySelector('[data-el="filters-toggle"]');

describe("toggleAdvancedFilters", () => {
  beforeEach(() => {
    document.body.innerHTML = markup;
  });

  it("clears inert on open and restores it on close", () => {
    expect(panel().inert).toBe(true);

    toggleAdvancedFilters();
    expect(panel().inert).toBe(false);
    expect(panel().classList.contains("visible")).toBe(true);

    toggleAdvancedFilters();
    expect(panel().inert).toBe(true);
    expect(panel().classList.contains("visible")).toBe(false);
  });

  it("mirrors the panel state on the toggle button", () => {
    toggleAdvancedFilters();
    expect(toggle().getAttribute("aria-expanded")).toBe("true");
    expect(toggle().classList.contains("active")).toBe(true);

    toggleAdvancedFilters();
    expect(toggle().getAttribute("aria-expanded")).toBe("false");
    expect(toggle().classList.contains("active")).toBe(false);
  });

  it("hands focus back to the toggle when closing from inside the panel", () => {
    toggleAdvancedFilters();
    const input = document.querySelector('[data-el="store-filter-input"]');
    input.focus();
    expect(document.activeElement).toBe(input);

    toggleAdvancedFilters();
    expect(document.activeElement).toBe(toggle());
  });

  it("leaves focus alone when closing from outside the panel", () => {
    toggleAdvancedFilters();
    const scrim = document.querySelector('[data-el="filter-sheet-scrim"]');
    scrim.focus();

    toggleAdvancedFilters();
    expect(document.activeElement).toBe(scrim);
  });
});
