/**
 * Client-side login gate.
 *
 * The app shell is served publicly and only the API routes are protected, so
 * this module asks the server whether a session exists, renders a login view
 * when it does not, and exposes a sign-out helper.
 *
 * The form is built through the DOM API rather than a template: the CSP is
 * `script-src 'self'` / `style-src 'self'` without `unsafe-inline`, so inline
 * handlers and `style` attributes are both blocked. Setting `el.style.*` from
 * script is unaffected — the same reasoning as `applyCategoryBadge()` in dom.js.
 */

const EYE_OPEN =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>';
const EYE_OFF =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c6.5 0 10 7 10 7a13.16 13.16 0 0 1-1.67 2.68"/><path d="M6.61 6.61A13.526 13.526 0 0 0 2 12s3.5 7 10 7a9.74 9.74 0 0 0 5.39-1.61"/><path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="2" y1="2" x2="22" y2="22"/></svg>';

/**
 * Fetch the current authentication status.
 *
 * @returns {Promise<{authed: boolean, enabled: boolean}>} `authed` is whether
 *   the app may start, `enabled` whether the password gate exists at all (which
 *   is what decides if the sign-out control is shown).
 */
export async function getAuthStatus() {
  try {
    const response = await fetch("/api/auth/me");
    // The service worker answers /api/* with a synthetic 503 while offline.
    // Treating that as "not logged in" would show a login form that cannot
    // possibly succeed, so an offline PWA is let through instead.
    if (response.status === 503) return { authed: true, enabled: false };
    if (!response.ok) return { authed: false, enabled: true };
    const data = await response.json();
    return { authed: data.authed === true, enabled: data.enabled === true };
  } catch {
    // Network error without a service worker: same reasoning as the 503 above.
    return { authed: true, enabled: false };
  }
}

/**
 * Sign out by clearing the server-side session cookie.
 *
 * @returns {Promise<void>}
 */
export async function logout() {
  try {
    await fetch("/api/auth/logout", { method: "POST" });
  } catch {
    // Ignore network errors: the cookie may already be gone, and the caller
    // reloads either way.
  }
}

/**
 * Reveal and wire the sign-out control.
 *
 * @param {boolean} enabled - Whether the deployment has the password gate on.
 *   When it does not, the button stays hidden rather than offering to end a
 *   session that does not exist.
 */
export function setupSignOut(enabled) {
  const button = document.querySelector('[data-el="logout"]');
  if (!button || !enabled) return;

  button.hidden = false;
  button.addEventListener("click", async () => {
    await logout();
    // A full reload rather than re-rendering the gate in place: signing out
    // ends the session deliberately, and reloading is the cheapest way to
    // guarantee every listener, timer and cached render is gone.
    location.reload();
  });
}

/**
 * Build the password field: an input plus a press-and-hold reveal toggle.
 *
 * @returns {{wrap: HTMLDivElement, input: HTMLInputElement}}
 */
function buildPasswordField() {
  const wrap = document.createElement("div");
  wrap.className = "login-password";

  const input = document.createElement("input");
  input.type = "password";
  input.className = "login-input";
  input.placeholder = "Password";
  // Without this, most password managers ignore the field entirely.
  input.autocomplete = "current-password";
  input.required = true;

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "login-reveal";
  toggle.setAttribute("aria-label", "Show password (press and hold)");
  toggle.innerHTML = EYE_OPEN;

  const show = () => {
    input.type = "text";
    toggle.innerHTML = EYE_OFF;
  };
  const hide = () => {
    input.type = "password";
    toggle.innerHTML = EYE_OPEN;
  };

  // Keep focus on the input, so Enter still submits the form instead of
  // re-triggering this button.
  toggle.addEventListener("mousedown", (event) => event.preventDefault());

  // Press-and-hold rather than a latch, and four ways to hide: a pointer can
  // leave the button, be cancelled by a gesture, or the button can lose focus —
  // any of which would otherwise strand the password in plain text.
  toggle.addEventListener("pointerdown", show);
  for (const name of ["pointerup", "pointerleave", "pointercancel", "blur"]) {
    toggle.addEventListener(name, hide);
  }
  toggle.addEventListener("keydown", (event) => {
    if (event.key === " " || event.key === "Enter") {
      event.preventDefault();
      show();
    }
  });
  toggle.addEventListener("keyup", (event) => {
    if (event.key === " " || event.key === "Enter") hide();
  });

  wrap.append(input, toggle);
  return { wrap, input };
}

/**
 * Build the "stay signed in" row.
 *
 * @returns {{label: HTMLLabelElement, checkbox: HTMLInputElement}}
 */
function buildRememberField() {
  const label = document.createElement("label");
  label.className = "login-remember";

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";

  const text = document.createElement("span");
  text.textContent = "Stay signed in for 30 days";

  label.append(checkbox, text);
  return { label, checkbox };
}

/**
 * Render the login view, hiding the app until the password is accepted.
 *
 * Safe to call again mid-session (an expired cookie): an existing gate is
 * replaced rather than stacked on top of.
 *
 * @param {() => void} onSuccess - Called once, after a successful login.
 */
export function renderLoginView(onSuccess) {
  // Modals are siblings of `.app`, not children, so both have to go. Hidden via
  // `style.display` rather than a class so nothing can collide with the app's
  // own display rules; restored to "" (not "block") to hand control back.
  const covered = Array.from(document.querySelectorAll(".app, .modal-overlay"));
  for (const element of covered) element.style.display = "none";

  document.querySelector('[data-el="login-gate"]')?.remove();

  const gate = document.createElement("div");
  gate.dataset.el = "login-gate";
  gate.className = "login-gate";

  const card = document.createElement("form");
  card.className = "login-card";
  // The form element buys Enter-to-submit and password-manager integration;
  // novalidate suppresses the native bubble so the inline message is the only
  // error UI, while `required` stays on the input as an accessibility hint.
  card.noValidate = true;

  const title = document.createElement("h1");
  title.className = "login-title";
  title.textContent = "Summa";

  const subtitle = document.createElement("p");
  subtitle.className = "login-subtitle";
  subtitle.textContent = "Enter your password to continue";

  const { wrap: passwordWrap, input: passwordInput } = buildPasswordField();
  const { label: rememberLabel, checkbox: rememberCheckbox } =
    buildRememberField();

  const error = document.createElement("p");
  error.className = "login-error";
  error.hidden = true;

  const submit = document.createElement("button");
  submit.type = "submit";
  submit.className = "login-submit";
  submit.textContent = "Sign in";

  card.append(title, subtitle, passwordWrap, rememberLabel, error, submit);
  gate.appendChild(card);
  document.body.appendChild(gate);
  passwordInput.focus();

  const fail = (message) => {
    error.textContent = message;
    error.hidden = false;
  };

  card.addEventListener("submit", async (event) => {
    event.preventDefault();
    error.hidden = true;
    submit.disabled = true;
    submit.textContent = "Signing in…";

    try {
      const response = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          password: passwordInput.value,
          remember: rememberCheckbox.checked,
        }),
      });

      if (response.status === 429) {
        fail("Too many attempts — please wait a few minutes");
        return;
      }
      if (!response.ok) {
        fail("Invalid password");
        // Pre-select the wrong password so retyping overwrites it.
        passwordInput.select();
        return;
      }

      gate.remove();
      for (const element of covered) element.style.display = "";
      onSuccess();
    } catch {
      fail("Network error — please try again");
    } finally {
      submit.disabled = false;
      submit.textContent = "Sign in";
    }
  });
}
