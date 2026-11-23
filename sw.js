/* SAFINAH · service worker
 * Buka lebih pantas & masih boleh dibuka tanpa internet (paparan terakhir).
 * Data sentiasa diambil terus dari Google Sheet — fail ini hanya simpan halaman app. */
const CACHE = 'safinah-app-v1';
self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './manifest.webmanifest'])).catch(() => {}));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== self.location.origin) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    try {
      const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), 4000);
      const res = await fetch(r, { signal: ctl.signal, cache: 'no-cache' });
      clearTimeout(t);
      if (res && res.ok) c.put(r.mode === 'navigate' ? './' : r, res.clone());
      return res;
    } catch (err) {
      const hit = (r.mode === 'navigate' ? await c.match('./') : null) || await c.match(r, { ignoreSearch: true });
      if (hit) return hit;
      throw err;
    }
  })());
});
