/**
 * Model picker for the AI Category Suggestions modal.
 *
 * The chosen model persists per-device in `localStorage` and is read by
 * `categorize.js` when it requests suggestions. Changing the model applies to
 * future runs only — it never re-runs the currently open modal.
 */

const STORAGE_KEY = "summa.aiModel";
const DEFAULT_KEY = "haiku";

// The single source of truth for the picker: option copy, trigger label, and
// the set of valid stored keys all derive from this list.
const MODELS = [
  {
    key: "haiku",
    name: "Claude Haiku",
    description: "Fast & low-cost · default",
  },
  {
    key: "sonnet",
    name: "Claude Sonnet",
    description: "More accurate on ambiguous stores",
  },
  { key: "opus", name: "Claude Opus", description: "Best quality · slowest" },
];

/**
 * Return the stored model key, validated against the known models. Falls back
 * to the default for a missing or unrecognized value.
 */
export function getAiModel() {
  const stored = localStorage.getItem(STORAGE_KEY);
  return MODELS.some((model) => model.key === stored) ? stored : DEFAULT_KEY;
}

function modelByKey(key) {
  return MODELS.find((model) => model.key === key) || MODELS[0];
}

/** Wire the header trigger and its dropdown menu. */
export function setupModelPicker() {
  const root = document.querySelector('[data-el="model-picker"]');
  if (!root) return;

  const trigger = root.querySelector('[data-el="model-picker-trigger"]');
  const label = root.querySelector('[data-el="model-picker-label"]');
  const menu = root.querySelector('[data-el="model-picker-menu"]');

  const syncSelection = () => {
    const current = getAiModel();
    label.textContent = modelByKey(current).name;
    menu.querySelectorAll("[data-model-option]").forEach((option) => {
      option.classList.toggle(
        "is-selected",
        option.dataset.modelOption === current,
      );
    });
  };

  const closeMenu = () => {
    if (!root.classList.contains("is-open")) return;
    root.classList.remove("is-open");
    trigger.setAttribute("aria-expanded", "false");
    document.removeEventListener("pointerdown", onOutside, true);
  };

  const onOutside = (event) => {
    if (!root.contains(event.target)) closeMenu();
  };

  const openMenu = () => {
    root.classList.add("is-open");
    trigger.setAttribute("aria-expanded", "true");
    document.addEventListener("pointerdown", onOutside, true);
  };

  trigger.addEventListener("click", () => {
    if (root.classList.contains("is-open")) closeMenu();
    else openMenu();
  });

  menu.addEventListener("click", (event) => {
    const option = event.target.closest("[data-model-option]");
    if (!option) return;
    localStorage.setItem(STORAGE_KEY, option.dataset.modelOption);
    syncSelection();
    closeMenu();
  });

  // Consume Esc while open so the first Esc closes the picker rather than the
  // whole modal (keyboard.js runs a global modal-stack Esc handler).
  root.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && root.classList.contains("is-open")) {
      event.stopPropagation();
      event.preventDefault();
      closeMenu();
      trigger.focus();
    }
  });

  syncSelection();
}
