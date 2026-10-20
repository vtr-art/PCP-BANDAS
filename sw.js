// PCP Bandas · service worker
// Troque o CACHE_NAME a cada nova versão para forçar os celulares a atualizarem.
const CACHE_NAME = 'pcp-bandas-v5';

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(['./', './index.html'])).catch(() => {}));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Sempre busca a versão nova na internet; usa a cópia guardada só se estiver sem conexão.
// Só mexe em arquivos do próprio site (mesma origem), nunca no Firebase ou no Google.
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  if (new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    fetch(req)
      .then(res => { const cp = res.clone(); caches.open(CACHE_NAME).then(c => c.put(req, cp)); return res; })
      .catch(() => caches.match(req).then(m => m || caches.match('./index.html')))
  );
});
