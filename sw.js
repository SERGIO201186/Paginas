const CACHE = 'pc-v3';
const ASSETS = [
  '/Paginas/login.html',
  '/Paginas/master.html',
  '/Paginas/admin/panel.html',
  'https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS).catch(()=>{}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // Solo cachear recursos locales, no el API de Apps Script
  if (e.request.url.includes('script.google.com')) return;
  // Red primero: así cada cambio publicado se ve de inmediato la próxima vez
  // que haya conexión. El cache queda solo como respaldo para cuando el
  // panel se abre sin internet (uso como PWA offline).
  e.respondWith(
    fetch(e.request).then(resp => {
      const copia = resp.clone();
      caches.open(CACHE).then(c => c.put(e.request, copia)).catch(()=>{});
      return resp;
    }).catch(() => caches.match(e.request))
  );
});
