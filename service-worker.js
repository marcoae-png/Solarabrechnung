const CACHE_NAME = 'solar-abrechnung-v1';
const ASSETS = [
  './',
  './index.html',
  './favicon.png',
  './manifest.json'
];

// Installation des Service Workers und Caching der Basisdaten
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS);
    })
  );
});

// Netzwerk-Anfragen abfangen (Lädt App-Struktur blitzschnell aus dem Cache)
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});
