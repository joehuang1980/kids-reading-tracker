// Service Worker — PWA offline support (Stage 15–16)
const CACHE = 'reading-tracker-v1';
const STATIC = [
  './',
  './index.html',
  './history.html',
  './dashboard.html',
  './css/style.css',
  './js/config.js',
  './js/isbn.js',
  './js/api.js',
  './js/scanner.js',
  './js/app.js',
  './js/history.js',
  './js/dashboard.js',
  './manifest.json',
  './assets/icons/icon.svg',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(STATIC)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Never intercept API calls — let them fail naturally for offline detection
  if (url.hostname.includes('googleapis.com') || url.hostname.includes('script.google.com')) {
    return;
  }

  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(resp => {
        if (resp.ok && e.request.method === 'GET') {
          const clone = resp.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return resp;
      });
    })
  );
});
