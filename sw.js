// Service Worker — Dashboard Mejor Riojanas
// Actualizar CACHE_VER fuerza limpieza en todos los browsers al próximo deploy
const CACHE_VER = '37cd632';

self.addEventListener('install', () => self.skipWaiting());

self.addEventListener('activate', event => {
  // Eliminar todos los caches anteriores
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Sin caché — siempre red para tener la versión más reciente
self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request, { cache: 'no-store' }).catch(() => fetch(event.request))
  );
});
