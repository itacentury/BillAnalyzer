import { describe, expect, it } from "vitest";

import { isTypingContext } from "../../static/js/keyboard.js";
import { isEditable } from "../../static/js/viewport.js";
import { highlightMatch } from "../../static/js/combobox.js";
import { itemRowInnerHtml } from "../../static/js/modals.js";

const el = (tag) => document.createElement(tag);

describe("isTypingContext", () => {
  it("is true for text-entry controls (including SELECT)", () => {
    expect(isTypingContext(el("input"))).toBe(true);
    expect(isTypingContext(el("textarea"))).toBe(true);
    expect(isTypingContext(el("select"))).toBe(true);
  });

  it("is false for a non-field element and a non-element target", () => {
    expect(isTypingContext(el("div"))).toBe(false);
    expect(isTypingContext(null)).toBe(false);
  });
});

describe("isEditable", () => {
  // Unlike isTypingContext, a SELECT does not raise the on-screen keyboard.
  it("is true for input and textarea but not select", () => {
    expect(isEditable(el("input"))).toBe(true);
    expect(isEditable(el("textarea"))).toBe(true);
    expect(isEditable(el("select"))).toBe(false);
  });

  it("is false for a plain element and a non-element target", () => {
    expect(isEditable(el("div"))).toBe(false);
    expect(isEditable(undefined)).toBe(false);
  });
});

describe("highlightMatch", () => {
  it("escapes the whole text when there is no query", () => {
    expect(highlightMatch("A & B", "")).toBe("A &amp; B");
  });

  it("wraps the first case-insensitive match in <b>", () => {
    expect(highlightMatch("Aldi", "ld")).toBe("A<b>ld</b>i");
    expect(highlightMatch("Aldi", "AL")).toBe("<b>Al</b>di");
  });

  it("escapes the text around a match so it can't inject markup", () => {
    expect(highlightMatch("<x>hit", "hit")).toBe("&lt;x&gt;<b>hit</b>");
  });

  it("returns the escaped text when the query does not match", () => {
    expect(highlightMatch("Aldi", "z")).toBe("Aldi");
  });
});

describe("itemRowInnerHtml", () => {
  it("renders empty inputs when no item is given", () => {
    const html = itemRowInnerHtml();

    expect(html).toContain('class="form-input item-name"');
    expect(html).not.toContain("value=");
  });

  it("pre-fills and escapes the inputs for an existing item", () => {
    const html = itemRowInnerHtml({ item_name: "A & B", item_price: "3.5" });

    expect(html).toContain('value="A &amp; B"');
    expect(html).toContain('value="3.5"');
  });
});
