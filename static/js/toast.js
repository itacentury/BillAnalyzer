/**
 * Toast notification system: an undo toast with a countdown, a plain notice
 * toast (success/info, no action), and an error toast for failures.
 *
 * Leaf module: resolves its DOM hooks lazily and imports only timing constants,
 * so it can be imported anywhere without creating cycles.
 */

import {
  UNDO_WINDOW_MS,
  NOTICE_WINDOW_MS,
  ERROR_TOAST_DURATION_MS,
  TOAST_SWAP_ANIMATION_MS,
} from "./timing.js";

let cachedEls = null;

/**
 * Return the cached toast element references, resolved from their `data-el`
 * hooks on first call (not at import time) so evaluation never depends on the
 * DOM already being parsed. Mirrors the `els()` pattern in dom.js.
 */
function toastEls() {
  if (!cachedEls) {
    const byHook = (name) => document.querySelector(`[data-el="${name}"]`);
    cachedEls = {
      undoToast: byHook("undoToast"),
      toastMessage: byHook("toastMessage"),
      toastUndo: byHook("toastUndo"),
      toastClose: byHook("toastClose"),
      toastProgress: byHook("toastProgress"),
      errorToast: byHook("errorToast"),
      errorToastMessage: byHook("errorToastMessage"),
      errorToastClose: byHook("errorToastClose"),
      errorToastProgress: byHook("errorToastProgress"),
    };
  }
  return cachedEls;
}

/**
 * Run the shrink-to-zero progress animation on `el` over `ms` milliseconds.
 * Resets to full width with transitions disabled, then re-enables the
 * transition on the next frame so the browser animates the collapse.
 */
function animateProgress(el, ms) {
  el.style.opacity = "1";
  el.style.transition = "none";
  el.style.transform = "scaleX(1)";
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      el.style.transition = `transform ${ms}ms linear`;
      el.style.transform = "scaleX(0)";
    });
  });
}

// Module-private toast state
let toastTimeout = null;
let toastUndoCallback = null;
// Called when the undo window closes without an undo (timeout, dismiss, being
// replaced by a newer toast, or page unload). Used to finalize a deferred
// action. Cleared once run so it never fires twice.
let toastCommitCallback = null;
let currentWindowMs = UNDO_WINDOW_MS;

/**
 * Run the pending commit callback exactly once, if any. Invoked whenever the
 * undo window resolves in favor of keeping the action (i.e. not undone).
 * Returns whether a commit actually ran, so a caller reloading the list can tell
 * an early finalize apart from a no-op re-entry and hide the now-stale toast.
 */
export function commitPendingToast() {
  const commit = toastCommitCallback;
  toastCommitCallback = null;
  if (commit) commit();
  return Boolean(commit);
}

/**
 * Show an undo toast with the given message.
 *
 * @param {string} message - Text to display in the toast.
 * @param {{onUndo?: Function, onCommit?: Function}} [options] -
 *   `onUndo` runs when the user clicks Undo; `onCommit` runs when the window
 *   closes without an undo (used to finalize a deferred action).
 */
export function showUndoToast(
  message,
  { onUndo = null, onCommit = null } = {},
) {
  presentToast(message, onUndo, onCommit, UNDO_WINDOW_MS);
}

/**
 * Show a transient notice toast (success/info) without an undo action.
 * Reuses the undo toast element with the Undo button hidden.
 */
export function showNoticeToast(message) {
  presentToast(message, null, null, NOTICE_WINDOW_MS);
}

/** Display or update the toast with a new message and callbacks. */
function presentToast(message, undoCallback, commitCallback, windowMs) {
  // A new toast replaces any pending one — finalize the outgoing action first
  // so it is never left unresolved.
  commitPendingToast();
  if (toastTimeout) {
    clearTimeout(toastTimeout);
    toastTimeout = null;
  }

  const { undoToast, toastMessage, toastUndo, toastProgress } = toastEls();
  toastUndoCallback = undoCallback;
  toastCommitCallback = commitCallback;
  currentWindowMs = windowMs;
  // Only offer the Undo button when there is something to undo.
  toastUndo.hidden = !undoCallback;
  const isVisible = undoToast.classList.contains("visible");

  const startCountdown = () => {
    animateProgress(toastProgress, currentWindowMs);
    toastTimeout = setTimeout(() => {
      commitPendingToast();
      hideUndoToast();
    }, currentWindowMs);
  };

  if (isVisible) {
    toastMessage.classList.add("swapping");
    setTimeout(() => {
      toastMessage.textContent = message;
      toastMessage.classList.remove("swapping");
      startCountdown();
    }, TOAST_SWAP_ANIMATION_MS);
  } else {
    undoToast.classList.add("visible");
    // Set the message only after the toast enters the a11y tree (it is
    // visibility:hidden until .visible) so the live region announces it.
    requestAnimationFrame(() => {
      toastMessage.textContent = message;
    });
    startCountdown();
  }
}

/**
 * Dismiss the toast and clear any pending timeout. Does not run the commit
 * callback — callers that mean "keep the action" call commitPendingToast() first.
 */
export function hideUndoToast() {
  if (toastTimeout) {
    clearTimeout(toastTimeout);
    toastTimeout = null;
  }
  toastUndoCallback = null;
  toastCommitCallback = null;
  toastEls().undoToast.classList.remove("visible");
}

/** Pause the toast countdown (e.g. on hover). */
function pauseToast() {
  if (toastTimeout) {
    clearTimeout(toastTimeout);
    toastTimeout = null;
  }
  toastEls().toastProgress.style.opacity = "0";
}

/** Resume the toast countdown after a pause. */
function resumeToast() {
  const { undoToast, toastProgress } = toastEls();
  if (!undoToast.classList.contains("visible")) return;
  animateProgress(toastProgress, currentWindowMs);
  toastTimeout = setTimeout(() => {
    commitPendingToast();
    hideUndoToast();
  }, currentWindowMs);
}

// Error toast state
let errorToastTimeout = null;

/**
 * Show an error toast with the given message. Auto-dismisses after a few seconds.
 *
 * @param {string} message - Text to display in the error toast.
 */
export function showErrorToast(message) {
  if (errorToastTimeout) {
    clearTimeout(errorToastTimeout);
    errorToastTimeout = null;
  }

  const { errorToast, errorToastMessage, errorToastProgress } = toastEls();
  errorToast.classList.add("visible");
  // Same reasoning as the undo toast: announce only once visible.
  requestAnimationFrame(() => {
    errorToastMessage.textContent = message;
  });

  animateProgress(errorToastProgress, ERROR_TOAST_DURATION_MS);

  errorToastTimeout = setTimeout(
    () => hideErrorToast(),
    ERROR_TOAST_DURATION_MS,
  );
}

/** Pause the error-toast countdown (e.g. on hover). */
function pauseErrorToast() {
  if (errorToastTimeout) {
    clearTimeout(errorToastTimeout);
    errorToastTimeout = null;
  }
  toastEls().errorToastProgress.style.opacity = "0";
}

/** Resume the error-toast countdown after a pause. */
function resumeErrorToast() {
  const { errorToast, errorToastProgress } = toastEls();
  if (!errorToast.classList.contains("visible")) return;
  animateProgress(errorToastProgress, ERROR_TOAST_DURATION_MS);
  errorToastTimeout = setTimeout(
    () => hideErrorToast(),
    ERROR_TOAST_DURATION_MS,
  );
}

/** Dismiss the error toast and clear any pending timeout. */
export function hideErrorToast() {
  if (errorToastTimeout) {
    clearTimeout(errorToastTimeout);
    errorToastTimeout = null;
  }
  toastEls().errorToast.classList.remove("visible");
}

/**
 * Attach toast interaction listeners (hover pause, undo/close buttons).
 * Must be called once during setup rather than at module load.
 */
export function initToastListeners() {
  const { undoToast, toastUndo, toastClose, errorToast, errorToastClose } =
    toastEls();

  undoToast.addEventListener("mouseenter", pauseToast);
  undoToast.addEventListener("mouseleave", resumeToast);

  errorToast.addEventListener("mouseenter", pauseErrorToast);
  errorToast.addEventListener("mouseleave", resumeErrorToast);

  toastUndo.addEventListener("click", async () => {
    // Undo cancels the deferred action — never commit it.
    toastCommitCallback = null;
    if (toastUndoCallback) await toastUndoCallback();
    hideUndoToast();
  });

  // Dismissing the toast accepts the action (e.g. finalizes the deletion).
  toastClose.addEventListener("click", () => {
    commitPendingToast();
    hideUndoToast();
  });
  errorToastClose.addEventListener("click", () => hideErrorToast());

  // If the page is closed while an undo window is open, finalize the pending
  // action (best-effort; the request may not complete during unload).
  window.addEventListener("beforeunload", commitPendingToast);
}
