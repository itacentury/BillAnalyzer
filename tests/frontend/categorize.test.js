/**
 * Frontend unit tests for the AI-categorize trigger badge. The badge must count
 * only the uncategorized invoices on the current page (`state.invoices`), so it
 * matches exactly what the AI action analyzes.
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { state } from "../../static/js/state.js";
import {
  runAnalysis,
  updateAiTriggerBadge,
} from "../../static/js/categorize.js";

function mountTriggerFixture() {
  document.body.innerHTML = `
    <button data-el="ai-categories-trigger">
      <span data-el="ai-categories-badge"></span>
    </button>
  `;
}

// runAnalysis writes the subtitle and (with an empty result) the content/footer,
// so the modal shell those selectors live in must be present.
function mountModalFixture() {
  document.body.innerHTML = `
    <div data-el="categorize-modal" class="active">
      <p data-el="categorize-subtitle"></p>
      <div data-el="categorize-content"></div>
      <div data-el="categorize-footer"></div>
    </div>
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

describe("runAnalysis", () => {
  beforeEach(() => {
    mountModalFixture();
    state.invoices = [];
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("posts only the ids of the uncategorized invoices on the current page", async () => {
    state.invoices = [
      { id: 1, category: null },
      { id: 2, category: "Groceries" },
      { id: 3, category: null },
    ];

    const jsonFor = (url) =>
      url.startsWith("/api/invoices/categorize-suggest")
        ? { suggestions: [], total: 0, count: 0 }
        : [];
    const fetchMock = vi.fn((url) =>
      Promise.resolve({
        ok: true,
        status: 200,
        json: () => Promise.resolve(jsonFor(url)),
      }),
    );
    vi.stubGlobal("fetch", fetchMock);

    await runAnalysis();

    const suggestCall = fetchMock.mock.calls.find(([url]) =>
      url.startsWith("/api/invoices/categorize-suggest"),
    );
    expect(JSON.parse(suggestCall[1].body).ids).toEqual([1, 3]);
  });
});
