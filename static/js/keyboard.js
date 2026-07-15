/**
 * Global keyboard shortcuts, modal focus-trap and the shortcut-help overlay.
 *
 * Modal focus behaviour (initial focus, Tab trapping, focus restoration) is
 * driven by a MutationObserver on the `.active` class of every `.modal-overlay`,
 * so the scattered open/close functions in modals.js / bulk.js stay untouched.
 */

import {
  openAddModal,
  openImportModal,
  lockScroll,
  unlockScroll,
} from "./modals.js";
import { els } from "./dom.js";

const FOCUSABLE_SELECTOR =
  'a[href], button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

// The element focused before the current modal opened, so focus can return
// there on close. Only one modal is ever active at a time, so a single ref
// suffices.
let lastTriggerElement = null;

/**
 * Whether the event target is a field where single-key shortcuts must not fire.
 */
function isTypingContext(target) {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName;
  return (
    tag === "INPUT" ||
    tag === "TEXTAREA" ||
    tag === "SELECT" ||
    target.isContentEditable
  );
}

function activeModal() {
  return document.querySelector(".modal-overlay.active");
}

/**
 * The tab-reachable, visible controls inside a container, in DOM order.
 * Excludes hidden inputs and `display:none` subtrees (no `offsetParent`) so the
 * focus trap never lands on the combobox's hidden value input or a hidden file
 * input.
 */
function getFocusable(container) {
  const nodes = [...container.querySelectorAll(FOCUSABLE_SELECTOR)];
  return nodes.filter((el) => el.type !== "hidden" && el.offsetParent !== null);
}

/**
 * Close a modal the way its own UI would: click its Cancel button (which also
 * resets the form). Overlays without one — the help overlay — just deactivate.
 */
function closeActiveModal(modal) {
  const cancelButton = modal.querySelector('[data-action="cancel"]');
  if (cancelButton) {
    cancelButton.click();
    return;
  }
  modal.classList.remove("active");
  unlockScroll();
}

function toggleShortcutHelp() {
  const overlay = document.querySelector('[data-el="shortcuts-help"]');
  if (!overlay) return;
  if (overlay.classList.toggle("active")) lockScroll();
  else unlockScroll();
}

/**
 * Keep Tab within the modal by wrapping focus around its first/last control.
 */
function trapTab(event, modal) {
  const focusables = getFocusable(modal);
  if (focusables.length === 0) return;

  const first = focusables[0];
  const last = focusables[focusables.length - 1];
  const active = document.activeElement;

  if (event.shiftKey && active === first) {
    event.preventDefault();
    last.focus();
  } else if (!event.shiftKey && active === last) {
    event.preventDefault();
    first.focus();
  }
}

function handleGlobalKeydown(event) {
  // `?` toggles the help overlay from anywhere outside a text field, but never
  // on top of a different modal.
  if (event.key === "?" && !isTypingContext(event.target)) {
    const modal = activeModal();
    const helpOverlay = document.querySelector('[data-el="shortcuts-help"]');
    if (!modal || modal === helpOverlay) {
      event.preventDefault();
      toggleShortcutHelp();
    }
    return;
  }

  const modal = activeModal();
  if (modal) {
    if (event.key === "Escape") {
      // Let a component that already handled Escape (an open combobox dropdown)
      // win: it closes first, and a second Escape then closes the modal.
      if (event.defaultPrevented) return;
      event.preventDefault();
      closeActiveModal(modal);
    } else if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      modal.querySelector(".modal-footer .btn-primary")?.click();
    } else if (event.key === "Tab") {
      trapTab(event, modal);
    }
    return;
  }

  // Global single-key shortcuts: only outside inputs and without modifiers.
  if (isTypingContext(event.target)) return;
  if (event.ctrlKey || event.metaKey || event.altKey) return;

  switch (event.key) {
    case "n":
      event.preventDefault();
      openAddModal();
      break;
    case "i":
      event.preventDefault();
      openImportModal();
      break;
    case "/":
      event.preventDefault();
      els().searchInput?.focus();
      break;
  }
}

function onModalOpen(overlay) {
  lastTriggerElement = document.activeElement;
  // Prefer the first field in the body over the header's close button.
  const body = overlay.querySelector(".modal-body");
  const target = (body && getFocusable(body)[0]) || getFocusable(overlay)[0];
  target?.focus();
}

function onModalClose() {
  if (lastTriggerElement instanceof HTMLElement) lastTriggerElement.focus();
  lastTriggerElement = null;
}

/**
 * Observe every modal overlay's `.active` toggle so focus is moved into the
 * modal on open and restored to the trigger on close, without touching the
 * open/close functions themselves.
 */
function setupModalFocusManagement() {
  const observer = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      const overlay = mutation.target;
      const isActive = overlay.classList.contains("active");
      const wasActive = (mutation.oldValue || "").includes("active");
      if (isActive && !wasActive) onModalOpen(overlay);
      else if (!isActive && wasActive) onModalClose();
    }
  });

  document.querySelectorAll(".modal-overlay").forEach((overlay) => {
    observer.observe(overlay, {
      attributes: true,
      attributeFilter: ["class"],
      attributeOldValue: true,
    });
  });
}

/**
 * Wire the global keydown shortcuts and modal focus management.
 */
export function setupKeyboardListeners() {
  document.addEventListener("keydown", handleGlobalKeydown);
  setupModalFocusManagement();
}
