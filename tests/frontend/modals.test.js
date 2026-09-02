/**
 * Frontend unit tests for the invoice date guard. The date input carries no
 * `max` attribute (Firefox for Android greys out the `max` day itself, making
 * today unpickable), so `validateInvoiceDate` is the only client-side check that
 * a future date is flagged — these cases cover it directly.
 */

import { beforeEach, describe, expect, it } from "vitest";

import { dateToIso, todayIso } from "../../static/js/dom.js";
import { validateInvoiceDate } from "../../static/js/modals.js";

// Built at local noon so the Berlin offset can't shift the calendar day.
function dayOffset(days) {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setDate(date.getDate() + days);
  return dateToIso(date);
}

function mountDateFixture() {
  document.body.innerHTML = `
    <input type="date" data-el="invoice-date" />
    <p class="field-error is-hidden" data-el="invoice-date-error"></p>
  `;
}

function dateInput() {
  return document.querySelector('[data-el="invoice-date"]');
}

function hintHidden() {
  return document
    .querySelector('[data-el="invoice-date-error"]')
    .classList.contains("is-hidden");
}

describe("validateInvoiceDate", () => {
  beforeEach(mountDateFixture);

  it("accepts today without flagging the field", () => {
    dateInput().value = todayIso();

    expect(validateInvoiceDate()).toBe(false);
    expect(hintHidden()).toBe(true);
    expect(dateInput().getAttribute("aria-invalid")).toBe("false");
  });

  it("flags a future date and shows the hint", () => {
    dateInput().value = dayOffset(1);

    expect(validateInvoiceDate()).toBe(true);
    expect(hintHidden()).toBe(false);
    expect(dateInput().getAttribute("aria-invalid")).toBe("true");
  });

  // Re-validating after a correction must clear the hint again, not leave the
  // form stuck in the invalid state.
  it("clears the hint once the date is corrected", () => {
    dateInput().value = dayOffset(1);
    validateInvoiceDate();

    dateInput().value = dayOffset(-1);

    expect(validateInvoiceDate()).toBe(false);
    expect(hintHidden()).toBe(true);
  });
});
