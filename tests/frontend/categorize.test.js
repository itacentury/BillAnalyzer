/**
 * Frontend unit tests for the AI-categorize trigger badge. The badge must count
 * only the uncategorized invoices on the current page (`state.invoices`), so it
 * matches exactly what the AI action analyzes.
 */

import { beforeEach, describe, expect, it } from "vitest";

import { state } from "../../static/js/state.js";
import { updateAiTriggerBadge } from "../../static/js/categorize.js";

function mountTriggerFixture() {
  document.body.innerHTML = `
    <button data-el="ai-categories-trigger">
      <span data-el="ai-categories-badge"></span>
    </button>
  `;
}

function button() {
  return document.querySelector('[data-el="ai-categories-trigger"]');
}

function badge() {
  return document.querySelector('[data-el="ai-categories-badge"]');
}

describe("updateAiTriggerBadge", () => {
  beforeEach(() => {
    mountTriggerFixture();
    state.invoices = [];
  });

  it("counts only the uncategorized invoices on the current page", () => {
    state.invoices = [
      { id: 1, category: null },
      { id: 2, category: "Groceries" },
      { id: 3, category: null },
    ];

    updateAiTriggerBadge();

    expect(badge().textContent).toBe("2");
    expect(button().disabled).toBe(false);
  });

  it("disables the trigger when the page has no uncategorized invoices", () => {
    state.invoices = [{ id: 1, category: "Groceries" }];

    updateAiTriggerBadge();

    expect(badge().textContent).toBe("0");
    expect(button().disabled).toBe(true);
  });
});
