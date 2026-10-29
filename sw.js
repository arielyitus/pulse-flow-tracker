const CACHE_VERSION = 'pulse-flow-hoops-v4.2';
const APP_SHELL_CACHE = `${CACHE_VERSION}-shell`;
const CDN_CACHE = `${CACHE_VERSION}-cdn`;

const PRECACHE_URLS = [
  './',
  './index.html',
  './manifest.json'
];

const CDN_HOSTS = [
  'cdn.tailwindcss.com',
  'cdn.jsdelivr.net',
  'unpkg.com',
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'www.gstatic.com'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(APP_SHELL_CACHE).then((cache) => {
      return cache.addAll(PRECACHE_URLS).catch(() => {});
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== APP_SHELL_CACHE && key !== CDN_CACHE)
          .map((key) => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);

  // Ignore Firebase Firestore / Auth live API traffic so Firestore IndexedDB handles offline sync
  if (
    url.hostname.includes('firestore.googleapis.com') ||
    url.hostname.includes('identitytoolkit.googleapis.com') ||
    url.hostname.includes('securetoken.googleapis.com')
  ) {
    return;
  }

  // 1. Cache-First strategy for CDN assets (Tailwind, Chart.js, Lucide, Fonts)
  if (CDN_HOSTS.some((host) => url.hostname.includes(host))) {
    event.respondWith(
      caches.open(CDN_CACHE).then(async (cache) => {
        const cached = await cache.match(request);
        const networkFetch = fetch(request)
          .then((response) => {
            if (response && (response.status === 200 || response.type === 'opaque')) {
              cache.put(request, response.clone());
            }
            return response;
          })
          .catch(() => cached);
        return cached || networkFetch;
      })
    );
    return;
  }

  // 2. Network-First strategy for index.html and same-origin app files
  if (url.origin === self.location.origin) {
    event.respondWith(
      fetch(request)
        .then((response) => {
          if (response && response.status === 200) {
            const copy = response.clone();
            caches.open(APP_SHELL_CACHE).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(async () => {
          const cached = await caches.match(request);
          if (cached) return cached;
          if (request.mode === 'navigate') {
            return caches.match('./index.html');
          }
          return new Response('Offline', { status: 503, statusText: 'Offline' });
        })
    );
  }
});
