/**
 * Shared animation/interaction timing constants (milliseconds).
 *
 * Kept in one leaf module so the JS durations and the CSS transition values in
 * toast.css stay in sync and are not scattered as magic numbers.
 */

// How long an undo toast stays open before its action is committed.
export const UNDO_WINDOW_MS = 5000;

// How long a plain notice (success/info, no undo) toast stays visible.
export const NOTICE_WINDOW_MS = 3000;

// Auto-dismiss delay for the error toast.
export const ERROR_TOAST_DURATION_MS = 4000;

// Message cross-fade when an open undo toast is replaced by a newer one.
// Mirrors the `.toast-message` opacity transition in toast.css.
export const TOAST_SWAP_ANIMATION_MS = 150;
