// Service worker: la app siempre intenta cargar la versión más reciente de tu
// página (red primero) y, si no hay internet, usa la copia guardada.
// Para forzar una limpieza de caché, cambia el número de VERSION.
const VERSION = 'v1';
const CACHE = `cronograma-${VERSION}`;

const ARCHIVOS = [
  './',
  './index.html',
  './css/style.css',
  './js/script.js',
  './favicon_io/favicon.ico',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.23/jspdf.plugin.autotable.min.js'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) =>
      // Si un archivo falla (p. ej. falta el favicon), no se cancela toda la instalación
      Promise.all(ARCHIVOS.map((u) => c.add(u).catch(() => {})))
    )
  );
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  const esMio = url.origin === self.location.origin;
  const esCDN = url.hostname === 'cdnjs.cloudflare.com';
  if (!esMio && !esCDN) return;

  if (esCDN) {
    // Librerías externas: caché primero (no cambian)
    e.respondWith(
      caches.match(e.request).then((hit) => hit || fetch(e.request).then((r) => {
        const copia = r.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copia));
        return r;
      }))
    );
    return;
  }

  // Tus archivos: red primero, caché como respaldo sin conexión
  e.respondWith(
    fetch(e.request).then((r) => {
      const copia = r.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copia));
      return r;
    }).catch(() => caches.match(e.request).then((hit) => hit || caches.match('./index.html')))
  );
});
