// Minimal service worker for PSX Dashboard — required by Chrome/Edge/Android
// for "Add to Home Screen" / Install installability checks.
const CACHE_NAME = 'psx-dashboard-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

// Pass-through fetch handler (required for installability on some browsers).
// Falls back to network; does not aggressively cache so the dashboard
// always reflects the latest data.
self.addEventListener('fetch', (event) => {
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});
