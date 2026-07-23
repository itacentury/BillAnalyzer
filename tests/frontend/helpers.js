/**
 * Shared fixtures for the frontend suite: a minimal `fetch` Response double and
 * the two DOM skeletons the API and render layers query via their `data-el`
 * hooks. `els()` (dom.js) caches on first call, so mount a fixture once per test
 * file and mutate values/state between cases rather than re-mounting.
 */

/**
 * A stand-in for a `fetch` Response carrying a JSON body. Only the members the
 * app touches (`ok`, `status`, `json()`) are provided.
 */
export function jsonResponse(data, { ok = true, status = 200 } = {}) {
  return { ok, status, json: async () => data };
}

/**
 * Mount the persistent filter inputs `els()` and `buildFilterParams()` read.
 * Every control is a plain input so `.value` is directly settable in a test.
 */
export function mountFilterFixture() {
  document.body.innerHTML = `
    <input data-el="search" />
    <input data-el="store-filter" />
    <input data-el="type-filter" />
    <input data-el="date-from" />
    <input data-el="date-to" />
    <input data-el="sort-by" value="date" />
    <input data-el="sort-order" value="desc" />
    <div data-el="month-display"></div>
    <div data-el="invoice-list"></div>
  `;
}

/**
 * Mount the containers `renderInvoices()` (and the pagination / bulk-toolbar
 * sub-renders it calls) write into.
 */
export function mountListFixture() {
  document.body.innerHTML = `
    <div data-el="invoice-list"></div>
    <span data-el="results-count"></span>
    <span data-el="results-total"></span>
    <div data-el="pagination"></div>
    <div data-el="bulk-action-toolbar">
      <button data-action="select-all"></button>
      <span data-el="selected-count"></span>
    </div>
    <label data-el="select-all-checkbox"><input type="checkbox" /></label>
  `;
}

/**
 * Yields to the next microtask so DOM updates from event handlers can settle.
 */
export function flushUi() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Mount the import modal containers and controls used by import UI tests.
 */
export function mountImportFixture() {
  document.body.innerHTML = `
    <div data-el="import-modal">
      <div class="modal-footer">
        <button type="button" class="btn btn-primary" data-action="import">Import</button>
      </div>
    </div>
    <div data-el="import-input"></div>
    <textarea data-el="json-input"></textarea>
    <div data-el="import-errors" class="is-hidden"></div>
    <div data-el="selected-files"></div>
    <div data-el="dropzone"></div>
    <input data-el="file-input" type="file" />
  `;
}
