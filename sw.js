/* خدمة بسيطة: بتخلي الموقع يتثبّت زي تطبيق ويفتح حتى لو النت ضعيف. البيانات نفسها بتتخزن في Firebase. */
const V = 'trucks-v1';
const SHELL = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL).catch(() => {})).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request; if (r.method !== 'GET') return;
  const u = new URL(r.url);
  const same = u.origin === location.origin;
  const lib = /(^|\.)(gstatic\.com|cdnjs\.cloudflare\.com|fonts\.googleapis\.com)$/.test(u.hostname) && !/firestore|identitytoolkit|securetoken/.test(u.href);
  if (!same && !lib) return;
  if (same) { /* الأحدث الأول، ولو مفيش نت يفتح من المحفوظ */
    e.respondWith(fetch(r).then(res => { const cp = res.clone(); caches.open(V).then(c => c.put(r, cp)); return res; }).catch(() => caches.match(r).then(m => m || caches.match('index.html'))));
  } else {
    e.respondWith(caches.match(r).then(m => m || fetch(r).then(res => { const cp = res.clone(); caches.open(V).then(c => c.put(r, cp)); return res; })));
  }
});
