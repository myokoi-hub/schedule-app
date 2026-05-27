const CACHE = 'schedule-app-v18';
// HTMLファイルはキャッシュしない（常に最新を取得するため）
const ASSETS = ['/style.css', '/manifest.json', '/icon-192.png', '/icon-512.png', '/schedule-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.url.includes('/api/')) return;

  // HTMLページは常にネットワークから取得（キャッシュバイパス）
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request, { cache: 'no-store' }).catch(() => caches.match(e.request)));
    return;
  }

  // CSS・画像等はネットワーク優先、失敗時はキャッシュから
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
