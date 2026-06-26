// Service Worker mínimo — solo para habilitar la instalación como PWA
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', () => self.clients.claim());
// Sin caché — el dashboard siempre carga desde la red para tener la versión más reciente
self.addEventListener('fetch', event => event.respondWith(fetch(event.request)));
