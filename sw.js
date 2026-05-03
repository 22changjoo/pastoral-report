const CACHE_NAME = 'shilmulga-v1';
const ASSETS = [
  '/pastoral-report/',
  '/pastoral-report/index.html',
  '/pastoral-report/manifest.json',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // 네트워크 우선, 실패시 캐시
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
