// Service worker minimal : requis pour l'installation en application.
// Les resultats viennent d'un service en ligne, rien n'est mis en cache.
const CACHE = 'jtcf-resultats-v1';
self.addEventListener('install', function (e) { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function (e) {
  e.respondWith(
    fetch(e.request).catch(function () { return caches.match(e.request); })
  );
});
