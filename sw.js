/* SAFINAH · service worker v2 (SAFINAH v3.7)
 * Buka serta-merta: halaman app diambil dari simpanan telefon dahulu, kemudian dikemas kini di belakang.
 * Data sentiasa diambil terus dari Google Sheet — fail ini hanya simpan halaman app. */
const CACHE = 'safinah-app-v2';
const AKAR = new URL('./', self.registration.scope).pathname;
const halApp = u => u.pathname === AKAR || u.pathname === AKAR + 'index.html';
self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './manifest.webmanifest'])).catch(() => {}));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
async function dariRangkaian(r, kunci, had) {
  const c = await caches.open(CACHE), ctl = new AbortController(), t = setTimeout(() => ctl.abort(), had);
  try {
    const res = await fetch(r, { signal: ctl.signal, cache: 'no-cache' });
    clearTimeout(t);
    if (res && res.ok) await c.put(kunci, res.clone());
    return res;
  } catch (err) {
    clearTimeout(t);
    const hit = await c.match(kunci, { ignoreSearch: true });
    if (hit) return hit;
    throw err;
  }
}
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== self.location.origin) return;
  if (r.mode === 'navigate') {
    e.respondWith((async () => {
      const hit = await (await caches.open(CACHE)).match('./');
      if (hit) { e.waitUntil(dariRangkaian(new Request('./'), './', 20000).catch(() => {})); return hit; }
      return dariRangkaian(r, './', 8000);
    })());
    return;
  }
  e.respondWith(dariRangkaian(r, halApp(u) ? './' : r, 4000));
});

