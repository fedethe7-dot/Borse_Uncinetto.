// Le borse di Mamma - service worker
// Quando aggiorni l'app, cambia il numero qui sotto (v1 -> v2) cosi' i telefoni scaricano la nuova versione.
const CACHE = 'borse-v1';
const FILES = ['./', 'index.html', 'manifest.json', 'icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  // Il foglio Google va sempre letto da internet
  if (url.hostname.includes('google.com') || url.hostname.includes('googleusercontent.com')) return;
  // Pagine dell'app: prima internet (per avere l'ultima versione), se manca la rete usa la copia salvata
  if (url.origin === location.origin) {
    e.respondWith(
      fetch(e.request).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r; })
        .catch(() => caches.match(e.request).then(r => r || caches.match('index.html')))
    );
    return;
  }
  // Caratteri (Google Fonts): copia salvata se c'e', altrimenti internet
  if (url.hostname.includes('fonts.g')) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return res; })));
  }
});
