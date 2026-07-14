/**
 * JSON file import flow: staging files, previewing them in the textarea and
 * sending them to the import endpoint.
 */

import { state } from "./state.js";
import { escapeHtml, showToast } from "./dom.js";
import { refreshAllData } from "./api.js";
import { closeImportModal } from "./modals.js";

export function handleMultipleFiles(files) {
  state.pendingFiles = Array.from(files).filter((f) =>
    f.name.endsWith(".json"),
  );
  updateSelectedFilesDisplay();

  // If files were selected, load them all into the textarea
  if (state.pendingFiles.length > 0) {
    loadFilesIntoTextarea();
  }
}

function updateSelectedFilesDisplay() {
  const container = document.querySelector('[data-el="selected-files"]');
  if (state.pendingFiles.length === 0) {
    container.style.display = "none";
    return;
  }

  container.style.display = "block";
  container.innerHTML = `
        <div style="background: var(--bg-tertiary); border-radius: var(--radius-sm); padding: 0.75rem;">
            <div style="font-size: 0.75rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.5rem;">
                ${state.pendingFiles.length} file(s) selected
            </div>
            ${state.pendingFiles
              .map(
                (f, i) => `
                <div style="display: flex; justify-content: space-between; align-items: center; padding: 0.375rem 0; border-bottom: 1px solid var(--border-subtle);">
                    <span style="font-size: 0.875rem;">📄 ${escapeHtml(
                      f.name,
                    )}</span>
                    <button type="button" class="btn btn-danger btn-sm" data-action="remove-file" data-index="${i}" style="padding: 0.25rem 0.5rem;">✕</button>
                </div>
            `,
              )
              .join("")}
        </div>
    `;
}

export function removeFile(index) {
  state.pendingFiles.splice(index, 1);
  updateSelectedFilesDisplay();
  loadFilesIntoTextarea();
}

/**
 * Wire the dropzone, the file picker and the staged-file remove buttons (the
 * latter via delegation, since the file list is rebuilt at runtime).
 */
export function setupImportListeners() {
  const dropzone = document.querySelector('[data-el="dropzone"]');
  const fileInput = document.querySelector('[data-el="file-input"]');

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
    const files = event.dataTransfer.files;
    if (files.length > 0) handleMultipleFiles(files);
  });
  fileInput.addEventListener("change", (event) => {
    if (event.target.files.length > 0) handleMultipleFiles(event.target.files);
  });

  document
    .querySelector('[data-el="selected-files"]')
    .addEventListener("click", (event) => {
      const removeButton = event.target.closest('[data-action="remove-file"]');
      if (removeButton) removeFile(Number(removeButton.dataset.index));
    });

  // Error-card actions are rebuilt at runtime, so delegate from the container.
  document
    .querySelector('[data-el="import-errors"]')
    .addEventListener("click", (event) => {
      if (event.target.closest('[data-action="reimport-corrected"]')) {
        reimportCorrected();
      } else if (event.target.closest('[data-action="download-invalid"]')) {
        downloadInvalid();
      }
    });
}

async function loadFilesIntoTextarea() {
  if (state.pendingFiles.length === 0) {
    document.querySelector('[data-el="json-input"]').value = "";
    return;
  }

  const allData = [];

  for (const file of state.pendingFiles) {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        allData.push(...parsed);
      } else {
        allData.push(parsed);
      }
    } catch {
      showToast(`Failed to read ${file.name}`, "error");
    }
  }

  document.querySelector('[data-el="json-input"]').value = JSON.stringify(
    allData,
    null,
    2,
  );
}

// Cap the inline editor cards so a huge failure set stays scannable; the rest
// are reachable via the download, which always exports every invalid entry.
const MAX_VISIBLE_ERRORS = 20;

export async function importJson() {
  const jsonText = document
    .querySelector('[data-el="json-input"]')
    .value.trim();
  if (!jsonText) {
    showToast("Please enter JSON data", "error");
    return;
  }

  let data;
  try {
    data = JSON.parse(jsonText);
    if (!Array.isArray(data)) data = [data];
  } catch {
    showToast("Invalid JSON format", "error");
    return;
  }

  await sendImport(data);
}

/**
 * POST invoice payloads to the import endpoint and report the outcome. Shared by
 * the initial import and the re-import of corrected entries: on partial success
 * the valid rows land and the failed entries are rendered as editable cards.
 */
async function sendImport(data) {
  const importButton = document.querySelector(
    '[data-el="import-modal"] .modal-footer .btn-primary',
  );
  const originalContent = importButton.innerHTML;
  importButton.innerHTML = '<div class="spinner"></div>';
  importButton.disabled = true;

  try {
    const response = await fetch("/api/invoices/import", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    });

    const result = await response.json();

    // A non-array payload (400) or a DB error (500) has no partial semantics.
    if (!response.ok || !result.success) {
      showToast(result.error || "Import failed", "error");
      return;
    }

    let message = `${result.imported} invoice(s) imported`;
    if (result.skipped > 0) {
      message += `, ${result.skipped} duplicate(s) skipped`;
    }
    if (result.failed > 0) {
      message += `, ${result.failed} failed`;
    }
    showToast(message, result.failed > 0 ? "error" : "success");

    // Valid entries always landed, so refresh the list regardless of failures.
    refreshAllData();

    if (result.failed > 0) {
      renderImportErrors(result.errors);
    } else {
      closeImportModal();
    }
  } catch {
    showToast("Import failed", "error");
  } finally {
    importButton.innerHTML = originalContent;
    importButton.disabled = false;
  }
}

/**
 * Render the invalid entries as per-entry editor cards. The full set is stashed on
 * state for download; only the first MAX_VISIBLE_ERRORS are shown as editable
 * cards, with a "+N more" note pointing at the download for the remainder.
 */
function renderImportErrors(errors) {
  state.importErrors = errors;
  const container = document.querySelector('[data-el="import-errors"]');

  const visible = errors.slice(0, MAX_VISIBLE_ERRORS);
  const hiddenCount = errors.length - visible.length;

  const cards = visible
    .map((error) => {
      const fieldLabel = error.field ? `${escapeHtml(error.field)}: ` : "";
      // Entities in the textarea body decode back to the raw JSON on parse, and
      // escaping prevents a value containing `</textarea>` from breaking out.
      const rawJson = escapeHtml(JSON.stringify(error.value, null, 2));
      return `
        <div class="import-error-card">
          <div class="import-error-head">
            <span class="import-error-index">#${error.index}</span>
            <span class="import-error-message">${fieldLabel}${escapeHtml(
              error.message,
            )}</span>
          </div>
          <textarea class="form-input error-editor" data-el="error-editor" data-index="${error.index}">${rawJson}</textarea>
        </div>
      `;
    })
    .join("");

  const moreNote =
    hiddenCount > 0
      ? `<div class="import-error-more">+${hiddenCount} more invalid entr${
          hiddenCount === 1 ? "y" : "ies"
        } — use Download to get them all</div>`
      : "";

  container.innerHTML = `
    <div class="import-error-title">${errors.length} entr${
      errors.length === 1 ? "y" : "ies"
    } failed — fix and re-import, or download</div>
    ${cards}
    ${moreNote}
    <div class="import-error-actions">
      <button type="button" class="btn btn-secondary btn-sm" data-action="download-invalid">Download invalid</button>
      <button type="button" class="btn btn-primary btn-sm" data-action="reimport-corrected">Re-import corrected</button>
    </div>
  `;
  container.style.display = "block";
}

/**
 * Collect the (possibly edited) JSON from every visible error card and re-import
 * only those entries — never the original full payload.
 */
async function reimportCorrected() {
  const editors = document.querySelectorAll('[data-el="error-editor"]');
  const data = [];
  for (const editor of editors) {
    try {
      data.push(JSON.parse(editor.value));
    } catch {
      showToast(`Entry #${editor.dataset.index}: invalid JSON`, "error");
      return;
    }
  }
  if (data.length === 0) return;
  await sendImport(data);
}

/**
 * Download every invalid entry (not just the visible cards) as a JSON file.
 */
function downloadInvalid() {
  if (state.importErrors.length === 0) return;
  const entries = state.importErrors.map((error) => error.value);
  const blob = new Blob([JSON.stringify(entries, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const today = new Date().toISOString().slice(0, 10);
  const link = document.createElement("a");
  link.href = url;
  link.download = `invalid-invoices-${today}.json`;
  link.click();
  URL.revokeObjectURL(url);
}
