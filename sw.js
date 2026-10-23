// Focus Tasks: the small helper that lets your phone install the page as an app.
// It only handles the page's own files (always fetching the newest copy, with the last copy
// as a fallback when you're offline). It never touches sign-in or your OneDrive data.
const CACHE = "focus-tasks-v4";
const FILES = ["./", "index.html", "app.js", "logic.js", "config.js", "msal-browser.min.js", "manifest.json", "icon-192.png", "icon-512.png"];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", event => {
  event.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return; // Microsoft and OneDrive: hands off
  event.respondWith(
    fetch(event.request)
      .then(response => {
        const copy = response.clone();
        if (response.ok) caches.open(CACHE).then(cache => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then(hit => hit || caches.match("index.html")))
  );
});
