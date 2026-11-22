/* Service worker – Thư viện số.
   - Giao diện (cùng tên miền): ưu tiên mạng, mất mạng thì mở bản đã lưu.
   - Thư viện quét mã / Excel trên cdn.jsdelivr.net: lưu lại sau lần tải đầu, lần sau mở ngay cả khi mạng yếu.
   - Dữ liệu Google Sheets (script.google.com) luôn lấy trực tiếp, không lưu. */
const CACHE = 'thuvien-v2';
const LIB = 'thuvien-lib-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];
const LIBS = ['https://cdn.jsdelivr.net/npm/@zxing/library@0.20.0/umd/index.min.js', 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'];
self.addEventListener('install', e => {
  e.waitUntil(Promise.all([
    caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}),
    caches.open(LIB).then(c => Promise.all(LIBS.map(u => c.match(u).then(r => r || fetch(u, { mode: 'cors' }).then(x => x.ok && c.put(u, x)).catch(() => {}))))).catch(() => {})
  ]));
  self.skipWaiting();
});
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== LIB).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const u = new URL(e.request.url);
  if (u.hostname === 'cdn.jsdelivr.net') {
    e.respondWith(caches.open(LIB).then(c => c.match(e.request.url).then(hit => hit || fetch(e.request).then(r => { if (r.ok || r.type === 'opaque') c.put(e.request.url, r.clone()); return r; }))));
    return;
  }
  if (u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request).then(r => r || caches.match('./index.html'))));
});
