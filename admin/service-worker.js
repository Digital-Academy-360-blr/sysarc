const CACHE_NAME = 'da360-admin-v3';
const CACHE_URLS = [
  './index.html',
  './manifest.json',
  'https://da360.vercel.app/images/logo.svg'
];

// Install - cache essential files
self.addEventListener('install', event => {
  console.log('[SW] Installing Service Worker v3');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[SW] Caching app shell');
        return cache.addAll(CACHE_URLS);
      })
      .catch(err => console.error('[SW] Cache failed:', err))
  );
  self.skipWaiting();
});

// Fetch - network first for API, cache first for static
self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  // Skip non-GET requests
  if (request.method !== 'GET') {
    return;
  }

  // CRITICAL: Never cache Google Apps Script API calls
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

  // Network-first strategy for HTML (always get fresh)
  if (request.mode === 'navigate' || request.destination === 'document') {
    event.respondWith(
      fetch(request)
        .then(response => {
          if (response.ok) {
            const responseClone = response.clone();
            caches.open(CACHE_NAME)
              .then(cache => cache.put(request, responseClone));
          }
          return response;
        })
        .catch(() => {
          return caches.match('./index.html')
            .then(cached => cached || new Response('Offline - Please check connection'));
        })
    );
    return;
  }

  // Cache-first for static assets
  event.respondWith(
    caches.match(request)
      .then(cached => {
        if (cached) {
          return cached;
        }
        return fetch(request)
          .then(response => {
            if (response.ok && response.status === 200) {
              const responseClone = response.clone();
              caches.open(CACHE_NAME)
                .then(cache => cache.put(request, responseClone));
            }
            return response;
          });
      })
      .catch(() => new Response('Network error'))
  );
});

// Activate - clean up old caches
self.addEventListener('activate', event => {
  console.log('[SW] Activating Service Worker v3');
  event.waitUntil(
    caches.keys()
      .then(cacheNames => {
        return Promise.all(
          cacheNames.map(cacheName => {
            if (cacheName !== CACHE_NAME) {
              console.log('[SW] Deleting old cache:', cacheName);
              return caches.delete(cacheName);
            }
          })
        );
      })
  );
  self.clients.claim();
});

// Message handler for cache refresh
self.addEventListener('message', event => {
  if (event.data === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  if (event.data === 'CLEAR_CACHE') {
    caches.delete(CACHE_NAME)
      .then(() => console.log('[SW] Cache cleared'));
  }
});