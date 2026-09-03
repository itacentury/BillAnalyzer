import { readFileSync } from "node:fs";
import { join } from "node:path";

import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  getAuthStatus,
  logout,
  renderLoginView,
  setupSignOut,
} from "../../static/js/auth.js";
import { jsonResponse } from "./helpers.js";

/**
 * The real component stylesheet, so the sign-out cases assert what a browser
 * would actually paint rather than only the `hidden` IDL property. Resolved via
 * `import.meta.dirname` because happy-dom replaces the global `URL`, and
 * `node:fs` rejects its file URLs.
 */
const sidebarCss = readFileSync(
  join(import.meta.dirname, "../../static/css/sidebar.css"),
  "utf8",
);

/** Submit the rendered login form and let its async handler settle. */
const submitLogin = async (password, { remember = false } = {}) => {
  const form = document.querySelector(".login-card");
  form.querySelector(".login-input").value = password;
  if (remember) form.querySelector(".login-remember input").checked = true;
  form.dispatchEvent(new Event("submit", { cancelable: true }));
  await vi.waitFor(() => expect(global.fetch).toHaveBeenCalled());
  // The handler awaits the response before touching the DOM.
  await Promise.resolve();
  await Promise.resolve();
};

const errorText = () =>
  document.querySelector(".login-error").hidden
    ? null
    : document.querySelector(".login-error").textContent;

beforeEach(() => {
  document.body.innerHTML = "";
  global.fetch = vi.fn();
});

describe("getAuthStatus", () => {
  it("reports the server's answer", async () => {
    global.fetch.mockResolvedValue(
      jsonResponse({ authed: true, enabled: true }),
    );

    await expect(getAuthStatus()).resolves.toEqual({
      authed: true,
      enabled: true,
    });
  });

  it("locks the app when the status request is refused", async () => {
    global.fetch.mockResolvedValue(
      jsonResponse({}, { ok: false, status: 500 }),
    );

    await expect(getAuthStatus()).resolves.toEqual({
      authed: false,
      enabled: true,
    });
  });

  it("lets an offline PWA boot when the service worker answers 503", async () => {
    // sw.js synthesizes a 503 for /api/* while offline. Showing a login form
    // there would be showing one that cannot possibly succeed.
    global.fetch.mockResolvedValue(
      jsonResponse({}, { ok: false, status: 503 }),
    );

    await expect(getAuthStatus()).resolves.toEqual({
      authed: true,
      enabled: false,
    });
  });

  it("lets the app boot when the network is gone entirely", async () => {
    global.fetch.mockRejectedValue(new Error("offline"));

    await expect(getAuthStatus()).resolves.toEqual({
      authed: true,
      enabled: false,
    });
  });

  it("treats a malformed payload as unauthenticated", async () => {
    global.fetch.mockResolvedValue(jsonResponse({ authed: "yes" }));

    await expect(getAuthStatus()).resolves.toEqual({
      authed: false,
      enabled: false,
    });
  });
});

describe("renderLoginView", () => {
  it("hides the app and the modals behind the gate", () => {
    document.body.innerHTML = `
      <div class="app"></div>
      <div class="modal-overlay"></div>
    `;

    renderLoginView(vi.fn());

    // Modals are siblings of .app, so hiding only .app would leave them on top.
    expect(document.querySelector(".app").style.display).toBe("none");
    expect(document.querySelector(".modal-overlay").style.display).toBe("none");
    expect(document.querySelector('[data-el="login-gate"]')).not.toBeNull();
  });

  it("replaces an existing gate instead of stacking a second one", () => {
    // Two concurrent expiries would otherwise leave one gate unreachable
    // underneath the other.
    renderLoginView(vi.fn());
    renderLoginView(vi.fn());

    expect(document.querySelectorAll('[data-el="login-gate"]')).toHaveLength(1);
  });

  it("sends the password and the remember flag", async () => {
    global.fetch.mockResolvedValue(jsonResponse({ authed: true }));
    renderLoginView(vi.fn());

    await submitLogin("hunter2", { remember: true });

    const [url, options] = global.fetch.mock.calls[0];
    expect(url).toBe("/api/auth/login");
    expect(options.method).toBe("POST");
    expect(JSON.parse(options.body)).toEqual({
      password: "hunter2",
      remember: true,
    });
  });

  it("restores the app and calls back once the password is accepted", async () => {
    document.body.innerHTML = '<div class="app"></div>';
    global.fetch.mockResolvedValue(jsonResponse({ authed: true }));
    const onSuccess = vi.fn();
    renderLoginView(onSuccess);

    await submitLogin("hunter2");

    expect(onSuccess).toHaveBeenCalledTimes(1);
    expect(document.querySelector('[data-el="login-gate"]')).toBeNull();
    expect(document.querySelector(".app").style.display).toBe("");
  });

  it("keeps the gate up and reports a wrong password", async () => {
    global.fetch.mockResolvedValue(
      jsonResponse({}, { ok: false, status: 401 }),
    );
    const onSuccess = vi.fn();
    renderLoginView(onSuccess);

    await submitLogin("wrong");

    expect(onSuccess).not.toHaveBeenCalled();
    expect(errorText()).toBe("Invalid password");
    expect(document.querySelector('[data-el="login-gate"]')).not.toBeNull();
  });

  it("distinguishes being throttled from being wrong", async () => {
    global.fetch.mockResolvedValue(
      jsonResponse({}, { ok: false, status: 429 }),
    );
    renderLoginView(vi.fn());

    await submitLogin("wrong");

    expect(errorText()).toMatch(/Too many attempts/);
  });

  it("reports a network failure without blaming the password", async () => {
    global.fetch.mockRejectedValue(new Error("offline"));
    renderLoginView(vi.fn());

    await submitLogin("hunter2");

    expect(errorText()).toMatch(/Network error/);
  });

  it("re-enables the submit button after a failure", async () => {
    global.fetch.mockResolvedValue(
      jsonResponse({}, { ok: false, status: 401 }),
    );
    renderLoginView(vi.fn());

    await submitLogin("wrong");

    expect(document.querySelector(".login-submit").disabled).toBe(false);
  });
});

describe("setupSignOut", () => {
  beforeEach(() => {
    // happy-dom ships no UA stylesheet, so the `[hidden]` default is stated
    // here. It comes first, exactly as the UA origin would: sidebar.css can
    // still override it by cascade order, which is the bug being guarded.
    document.body.innerHTML = `
      <style>[hidden] { display: none }</style>
      <style>${sidebarCss}</style>
      <button class="sidebar-action-btn" data-el="logout" hidden></button>`;
  });

  it("reveals the control when the gate is enabled", () => {
    setupSignOut(true);

    const button = document.querySelector('[data-el="logout"]');
    expect(button.hidden).toBe(false);
    expect(getComputedStyle(button).display).toBe("flex");
  });

  it("leaves it hidden on a deployment without a password", () => {
    setupSignOut(false);

    const button = document.querySelector('[data-el="logout"]');
    expect(button.hidden).toBe(true);
    expect(getComputedStyle(button).display).toBe("none");
  });
});

describe("logout", () => {
  it("posts to the logout endpoint", async () => {
    global.fetch.mockResolvedValue(jsonResponse({ authed: false }));

    await logout();

    expect(global.fetch).toHaveBeenCalledWith("/api/auth/logout", {
      method: "POST",
    });
  });

  it("resolves even when the request fails", async () => {
    // The caller reloads either way, and the cookie may already be gone.
    global.fetch.mockRejectedValue(new Error("offline"));

    await expect(logout()).resolves.toBeUndefined();
  });
});
