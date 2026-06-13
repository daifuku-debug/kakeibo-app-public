const CACHE_NAME = 'kakeibo-pwa-v29';
const urlsToCache = [
  './',
  './index.html',
  './style.css?v=quests-1',
  './app.js?v=quests-1',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(urlsToCache))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  if (event.request.method !== 'GET' || url.origin !== self.location.origin) {
    return;
  }

  const isFreshAppAsset = url.pathname.endsWith('/app.js') || url.pathname.endsWith('/style.css');
  if (event.request.mode === 'navigate' || isFreshAppAsset) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request.mode === 'navigate' ? './index.html' : event.request, copy));
          return response;
        })
        .catch(() => caches.match(event.request.mode === 'navigate' ? './index.html' : event.request))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(response => {
      const network = fetch(event.request)
        .then(fresh => {
          const copy = fresh.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
          return fresh;
        })
        .catch(() => response);
      return response || network;
    })
  );
});
