/**
 * JSON file import flow: staging files, previewing them in the textarea and
 * sending them to the import endpoint.
 */

import { state } from "./state.js";
import { escapeHtml, showToast } from "./dom.js";
import { loadInvoices, loadStores } from "./api.js";
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
                    <button type="button" class="btn btn-danger btn-sm" onclick="removeFile(${i})" style="padding: 0.25rem 0.5rem;">✕</button>
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
    if (result.success) {
      let message = `${result.imported} invoice(s) imported`;
      if (result.skipped > 0) {
        message += `, ${result.skipped} duplicate(s) skipped`;
      }
      showToast(message, "success");
      closeImportModal();
      loadInvoices();
      loadStores();
    } else {
      showToast("Import failed", "error");
    }
  } catch {
    showToast("Import failed", "error");
  } finally {
    importButton.innerHTML = originalContent;
    importButton.disabled = false;
  }
}
