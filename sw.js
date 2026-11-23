// Focus Tasks: the small helper that lets your phone install the page as an app.
// It only handles the page's own files (always fetching the newest copy, with the last copy
// as a fallback when you're offline). It never touches sign-in or your OneDrive data.
const CACHE = "focus-tasks-v12";
const FILES = ["./", "index.html", "boot.js?v=12", "app.js?v=12", "logic.js?v=12", "config.js?v=12", "msal-browser.min.js?v=12", "manifest.json", "icon-192.png", "icon-512.png",
  "training.css?v=12", "training-data.js?v=12", "training-reader.js?v=12", "training-ui.js?v=12",
  "fonts/literata-latin-wght-normal.woff2", "fonts/literata-latin-wght-italic.woff2", "fonts/literata-latin-ext-wght-normal.woff2", "fonts/literata-latin-ext-wght-italic.woff2",
  "fonts/literata-greek-wght-normal.woff2", "fonts/literata-greek-wght-italic.woff2", "fonts/source-serif-4-latin-wght-normal.woff2", "fonts/source-serif-4-latin-wght-italic.woff2",
  "fonts/source-serif-4-greek-wght-normal.woff2", "fonts/source-serif-4-greek-wght-italic.woff2", "fonts/atkinson-hyperlegible-latin-400-normal.woff2",
  "fonts/atkinson-hyperlegible-latin-400-italic.woff2", "fonts/atkinson-hyperlegible-latin-700-normal.woff2"];

self.addEventListener("install", event => {
  // One missing file (say, a font) never stops the update.
  event.waitUntil(caches.open(CACHE).then(cache => Promise.all(FILES.map(f => cache.add(f).catch(() => {})))).then(() => self.skipWaiting()));
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
    // Always ask for the newest copy (never an old one from the browser's own cache), so the files always match.
    fetch(event.request, { cache: "no-cache" })
      .then(response => {
        const copy = response.clone();
        if (response.ok) caches.open(CACHE).then(cache => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then(hit => hit || caches.match("index.html")))
  );
});
