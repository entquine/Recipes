// Service worker: keeps the app working offline.
const VERSION = 'wcwb-v-018';
const CACHE = 'dinner-' + VERSION;
const FILES = ["./","index.html","styles.css","app.js","config.js","manifest.webmanifest","recipes.xlsx","icons/icon-192.png","icons/icon-512.png","icons/icon-maskable-512.png"];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k.startsWith('dinner-') && k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
// Same-origin files: answer from cache straight away, refresh the cache in the background.
self.addEventListener('fetch', (e) => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  if (url.pathname.endsWith('/sw.js')) return; // always ask the network for the latest version
  const key = req.mode === 'navigate' ? './' : req;
  if (url.pathname.endsWith('/recipes.xlsx')) { // recipes: try the network first (4 s), so edits show up straight away
    e.respondWith(caches.open(CACHE).then(async (cache) => {
      const cached = await cache.match(key, { ignoreSearch: true });
      const net = fetch(req, { cache: 'no-cache' }).then((res) => { if (res.ok) cache.put(key, res.clone()); return res; });
      const timeout = new Promise((resolve) => setTimeout(() => resolve(null), 4000));
      const first = await Promise.race([net.catch(() => null), timeout]);
      if (first && first.ok) return first;
      if (cached) { e.waitUntil(net.catch(() => {})); return cached; }
      return net.catch(() => new Response('Offline', { status: 503 }));
    }));
    return;
  }
  e.respondWith(caches.open(CACHE).then(async (cache) => {
    const cached = await cache.match(key, { ignoreSearch: true });
    const fresh = fetch(req, { cache: 'no-cache' }).then((res) => {
      if (res.ok) cache.put(key, res.clone());
      return res;
    }).catch(() => cached);
    if (cached) { e.waitUntil(fresh); return cached; }
    return fresh.then((r) => r || new Response('Offline', { status: 503 }));
  }));
});
