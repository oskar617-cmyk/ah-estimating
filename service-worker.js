// AH Estimating service worker
// Minimal pass-through SW. We don't cache anything so the app is always fresh
// and never serves stale data from SharePoint. This is enough to make the app
// installable as a PWA.

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  // Pass through every request to the network.
  // No caching - we want fresh data from SharePoint and Microsoft Graph every time.
  return;
});
