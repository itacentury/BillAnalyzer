/**
 * Swipe-down-to-close gesture for the mobile bottom sheets (add/edit invoice,
 * JSON import, filter panel). The drag starts on a sheet's grabber or header,
 * follows the pointer downwards and either closes the sheet (past the
 * threshold) or snaps it back.
 */

import { mobileViewport } from "./dom.js";

const CLOSE_THRESHOLD_PX = 80;

/** Return the draggable sheet element for a grabber/header hit, if any. */
function sheetFromHandle(target) {
  const handle = target.closest(".sheet-grabber, .modal-header, .sheet-header");
  if (!handle) return null;
  return handle.closest(".modal-sheet, .filters-collapsible");
}

/** Trigger the sheet's own close control so state stays in one place. */
function closeSheet(sheet) {
  const closeButton = sheet.classList.contains("filters-collapsible")
    ? document.querySelector('[data-action="close-filter-sheet"]')
    : sheet.querySelector('[data-action="cancel"], .modal-close');
  if (closeButton) closeButton.click();
}

export function setupSheetGestures() {
  let sheet = null;
  let startY = 0;
  let deltaY = 0;

  document.addEventListener("pointerdown", (event) => {
    if (!mobileViewport.matches) return;
    sheet = sheetFromHandle(event.target);
    if (!sheet) return;
    startY = event.clientY;
    deltaY = 0;
    sheet.style.transition = "none";
    event.target.setPointerCapture(event.pointerId);
  });

  document.addEventListener("pointermove", (event) => {
    if (!sheet) return;
    deltaY = Math.max(0, event.clientY - startY);
    sheet.style.transform = `translateY(${deltaY}px)`;
  });

  const release = () => {
    if (!sheet) return;
    const dragged = sheet;
    sheet = null;
    dragged.style.transition = "";
    // Close first, clear the inline transform after: the inline value still
    // holds the dragged position when the close retargets the CSS transform,
    // so the filter sheet's close transition continues downward from the
    // finger instead of snapping back to 0 first.
    if (deltaY > CLOSE_THRESHOLD_PX) closeSheet(dragged);
    dragged.style.transform = "";
  };

  document.addEventListener("pointerup", release);
  document.addEventListener("pointercancel", release);
}
