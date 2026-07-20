/**
 * Keep the mobile bottom sheets clear of the on-screen keyboard.
 *
 * The sheets are `position: fixed` and sized in `vh`, both of which ignore the
 * virtual keyboard, so a focused field can end up hidden behind it. This module
 * mirrors the `visualViewport` (the region actually visible above the keyboard)
 * into two CSS custom properties the sheet CSS consumes, and scrolls the focused
 * field into the centre of that region.
 */

import { mobileViewport } from "./dom.js";

// How long to wait after focus before centring the field, so the keyboard has
// animated in and the visual viewport has resized first.
const SCROLL_SETTLE_MS = 300;

const SHEET_SELECTOR = ".modal-sheet, .filters-collapsible";

/**
 * Whether the element is a field that raises the on-screen keyboard.
 */
export function isEditable(element) {
  if (!(element instanceof HTMLElement)) return false;
  const tag = element.tagName;
  return tag === "INPUT" || tag === "TEXTAREA" || element.isContentEditable;
}

/**
 * Mirror the visual viewport into `--visual-viewport-height` (its visible
 * height) and `--keyboard-inset` (how far the keyboard overlaps the layout
 * viewport, used to lift a bottom-pinned sheet above it).
 */
function syncViewport(viewport) {
  const root = document.documentElement.style;
  root.setProperty("--visual-viewport-height", `${viewport.height}px`);

  // `visualViewport.height` also shrinks when the browser address bar appears,
  // so only treat it as keyboard overlap while a field is focused on mobile.
  const keyboardOpen =
    mobileViewport.matches && isEditable(document.activeElement);
  const overlap = window.innerHeight - (viewport.height + viewport.offsetTop);
  const inset = keyboardOpen ? Math.max(0, overlap) : 0;
  root.setProperty("--keyboard-inset", `${inset}px`);
}

/**
 * Wire the visual-viewport mirroring and focus-into-view behaviour. A no-op on
 * browsers without the `visualViewport` API — the CSS falls back to `dvh`.
 */
export function setupViewportListeners() {
  const viewport = window.visualViewport;
  if (!viewport) return;

  const onViewportChange = () => syncViewport(viewport);
  syncViewport(viewport);
  viewport.addEventListener("resize", onViewportChange);
  viewport.addEventListener("scroll", onViewportChange);

  document.addEventListener("focusin", (event) => {
    if (!mobileViewport.matches) return;
    const field = event.target;
    if (!isEditable(field) || !field.closest(SHEET_SELECTOR)) return;
    setTimeout(
      () => field.scrollIntoView({ block: "center" }),
      SCROLL_SETTLE_MS,
    );
  });
}
