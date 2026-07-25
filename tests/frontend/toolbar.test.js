import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { mountToolbarFixture } from "./helpers.js";

// The toolbar listeners reload the list on every interaction; the network layer
// is not what these tests are about.
vi.mock("../../static/js/api.js", () => ({
  loadInvoices: vi.fn(),
}));

// happy-dom has no layout engine, so `searchFieldSpace` is exercised with
// explicit measurements. These stand in for a roomy desktop toolbar.
const measurements = {
  rowWidth: 1200,
  wrapped: false,
  aiPresent: true,
  aiWidth: 40,
  filterWidth: 90,
  quickFloor: 200,
  monthWidth: 180,
  todayWidth: 60,
  rowGap: 8,
};

describe("searchFieldSpace", () => {
  let searchFieldSpace;

  beforeEach(async () => {
    ({ searchFieldSpace } = await import("../../static/js/filters.js"));
  });

  it("subtracts the fixed chips and the inline gaps", () => {
    // 1200 - (200 + 180 + 60 + 40 + 90) - 8 * 5
    expect(searchFieldSpace(measurements)).toBe(590);
  });

  it("counts one gap fewer when the AI trigger is absent", () => {
    // The regression this guards: dropping the element also drops one flex gap,
    // so keeping the AI-present gap count under-reports the room by exactly one
    // gap and leaves the field collapsed to an icon for no reason.
    const withAi = searchFieldSpace(measurements);
    const withoutAi = searchFieldSpace({
      ...measurements,
      aiPresent: false,
      aiWidth: 0,
    });
    expect(withoutAi - withAi).toBe(measurements.aiWidth + measurements.rowGap);
  });

  it("counts the AI gap from presence, not from measured width", () => {
    // A trigger measuring zero is one that is not laid out yet, not one that is
    // absent: the flex item — and therefore its gap — is still in the row.
    const unmeasured = searchFieldSpace({ ...measurements, aiWidth: 0 });
    const absent = searchFieldSpace({
      ...measurements,
      aiPresent: false,
      aiWidth: 0,
    });
    expect(absent - unmeasured).toBe(measurements.rowGap);
  });

  describe("wrapped onto its own row", () => {
    const wrapped = { ...measurements, wrapped: true };

    it("only subtracts the items sharing that second row", () => {
      // 1200 - 40 - 90 - 8 * 2
      expect(searchFieldSpace(wrapped)).toBe(1054);
    });

    it("counts one gap fewer without the AI trigger", () => {
      const withoutAi = searchFieldSpace({
        ...wrapped,
        aiPresent: false,
        aiWidth: 0,
      });
      expect(withoutAi - searchFieldSpace(wrapped)).toBe(
        wrapped.aiWidth + wrapped.rowGap,
      );
    });
  });
});

describe("setupFilterListeners with the AI trigger absent", () => {
  let observe;

  beforeEach(() => {
    observe = vi.fn((target) => {
      // Spec-faithful stub. happy-dom's own ResizeObserver accepts null without
      // complaint, so testing against it would pass even with the null-guard
      // deleted. A real browser throws here — that throw is what killed the rest
      // of init() when AI suggestions were disabled.
      if (!(target instanceof Element)) {
        throw new TypeError("parameter 1 is not of type 'Element'");
      }
    });
    vi.stubGlobal(
      "ResizeObserver",
      class {
        observe = observe;
        disconnect() {}
        unobserve() {}
      },
    );
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.resetModules();
  });

  // `els()` caches its lookups on first call, so each configuration needs a
  // fresh module registry rather than only a fresh DOM.
  const setupWith = async (options) => {
    vi.resetModules();
    mountToolbarFixture(options);
    const { setupFilterListeners } = await import("../../static/js/filters.js");
    setupFilterListeners();
  };

  it("wires up without throwing", async () => {
    await expect(setupWith({ aiTrigger: false })).resolves.not.toThrow();
  });

  it("observes only the toolbar row", async () => {
    await setupWith({ aiTrigger: false });
    expect(observe).toHaveBeenCalledTimes(1);
    expect(observe).toHaveBeenCalledWith(
      document.querySelector(".filters-row"),
    );
  });

  it("observes the trigger as well when it is present", async () => {
    await setupWith({ aiTrigger: true });
    expect(observe).toHaveBeenCalledTimes(2);
    expect(observe).toHaveBeenCalledWith(
      document.querySelector('[data-el="ai-categories-trigger"]'),
    );
  });

  it("runs the setup steps that come after the observer", async () => {
    // The point of the guard: a missing element must not cost the toolbar the
    // rest of its wiring. updateFilterBadge() is the last statement of
    // setupFilterListeners(), so a hidden badge means execution reached the end.
    await setupWith({ aiTrigger: false });
    const badge = document.querySelector('[data-el="filter-badge"]');
    expect(badge.hidden).toBe(true);
  });
});
