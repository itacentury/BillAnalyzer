/**
 * Service Worker for Summa PWA.
 * Provides offline support and caching strategies.
 */

const CACHE_NAME = "summa-cache-v7";
// JS modules are not listed here: they are discovered at install time from the
// server-rendered /static/js-manifest.json (which globs static/js/*.js), so
// adding a module needs no edit to this file.
const STATIC_ASSETS = [
  "/",
  "/static/css/style.css",
  "/static/favicon.svg",
  "/static/manifest.json",
  "/static/icons/icon-192.png",
  "/static/icons/icon-512.png",
];

/**
 * Install event - cache static assets.
 */
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      console.log("[SW] Caching static assets");
      try {
        const response = await fetch("/static/js-manifest.json");
        const jsModules = await response.json();
        await cache.addAll([...STATIC_ASSETS, ...jsModules]);
      } catch (error) {
        console.log(
          "[SW] JS manifest unavailable, caching core assets only:",
          error,
        );
        await cache.addAll(STATIC_ASSETS);
      }
    }),
  );
  // Activate immediately
  self.skipWaiting();
});

/**
 * Activate event - clean up old caches.
 */
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames
          .filter((name) => name !== CACHE_NAME)
          .map((name) => {
            console.log("[SW] Deleting old cache:", name);
            return caches.delete(name);
          }),
      );
    }),
  );
  // Take control of all pages immediately
  self.clients.claim();
});

/**
 * Fetch event - network-first strategy for API, cache-first for static assets.
 */
self.addEventListener("fetch", (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // Skip non-GET requests
  if (request.method !== "GET") {
    return;
  }

  // API requests: Network-first, fallback to cache
  if (url.pathname.startsWith("/api/")) {
    event.respondWith(networkFirst(request));
    return;
  }

  // Static assets: Cache-first, fallback to network
  event.respondWith(cacheFirst(request));
});

/**
 * Cache-first strategy: Try cache, fallback to network.
 */
async function cacheFirst(request) {
  const cachedResponse = await caches.match(request);
  if (cachedResponse) {
    return cachedResponse;
  }

  try {
    const networkResponse = await fetch(request);
    // Cache successful responses
    if (networkResponse.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, networkResponse.clone());
    }
    return networkResponse;
  } catch (error) {
    // Return offline fallback for navigation requests
    if (request.mode === "navigate") {
      return caches.match("/");
    }
    throw error;
  }
}

/**
 * Network-first strategy: Try network, fallback to cache.
 */
async function networkFirst(request) {
  try {
    const networkResponse = await fetch(request);
    // Cache successful GET responses
    if (networkResponse.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, networkResponse.clone());
    }
    return networkResponse;
  } catch {
    const cachedResponse = await caches.match(request);
    if (cachedResponse) {
      return cachedResponse;
    }
    // Return error response for failed API calls
    return new Response(JSON.stringify({ error: "Offline - no connection" }), {
      status: 503,
      headers: { "Content-Type": "application/json" },
    });
  }
}
