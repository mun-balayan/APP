// APP-CSE Service Worker — network-first
const CACHE = 'app-cse-v17';
const ASSETS = ['./', './index.html'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin === location.origin) {
    e.respondWith(
      fetch(e.request)
        .then(res => { if(res.ok){ const c=res.clone(); caches.open(CACHE).then(ca=>ca.put(e.request,c)); } return res; })
        .catch(() => caches.match(e.request))
    );
  }
});
