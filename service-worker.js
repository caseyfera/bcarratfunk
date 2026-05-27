// ============================================================
//  Dragonfly — Service Worker v2.96
//  Caches the app shell so it loads instantly and works offline
// ============================================================

const CACHE_NAME = 'dragonfly-v2.96';

// Files to cache on install — the core app shell
const PRECACHE = [
  '/dragonfly_v2_96.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
];

// ── Install: cache the app shell ────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(PRECACHE))
      .then(() => self.skipWaiting())
  );
});

// ── Activate: clean up old caches ───────────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// ── Fetch: serve from cache, fall back to network ───────────
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Always go network-first for the Apps Script proxy
  // (game data must be fresh — never serve stale api responses)
  if (url.pathname.startsWith('/api/gas')) {
    event.respondWith(fetch(event.request));
    return;
  }

  // Cache-first for everything else (app shell, icons, manifest)
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // Cache valid responses for future offline use
        if (response && response.status === 200 && response.type === 'basic') {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        }
        return response;
      });
    })
  );
});
