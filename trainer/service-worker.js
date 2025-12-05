const CACHE_NAME = 'da360-trainer-v3';
const CACHE_URLS = [
  './index.html',
  './manifest.json',
  'https://da360.vercel.app/images/logo.svg'
];

self.addEventListener('install', event => {
  console.log('[SW Trainer] Installing v3');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(CACHE_URLS))
      .catch(err => console.error('[SW Trainer] Cache failed:', err))
  );
  self.skipWaiting();
});

self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  if (request.method !== 'GET') {
    return;
  }

  // NEVER cache API calls - always fresh from Google Apps Script
  if (url.hostname.includes('script.google.com') || 
      url.hostname.includes('script.googleusercontent.com')) {
    event.respondWith(
      fetch(request, {
        cache: 'no-store',
        headers: {
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache'
        }
      })
    );
    return;
  }

  // Network-first for HTML pages
  if (request.mode === 'navigate' || request.destination === 'document') {
    event.respondWith(
      fetch(request)
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(request, clone));
          }
          return response;
        })
        .catch(() => caches.match('./index.html')
          .then(cached => cached || new Response('Offline')))
    );
    return;
  }

  // Cache-first for static resources
  event.respondWith(
    caches.match(request)
      .then(cached => cached || fetch(request)
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(request, clone));
          }
          return response;
        }))
      .catch(() => new Response('Network error'))
  );
});

self.addEventListener('activate', event => {
  console.log('[SW Trainer] Activating v3');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('message', event => {
  if (event.data === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  if (event.data === 'CLEAR_CACHE') {
    caches.delete(CACHE_NAME);
  }
});