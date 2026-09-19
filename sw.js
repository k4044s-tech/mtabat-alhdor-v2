const CACHE_NAME = 'hather-shell-v1';
const SHELL_FILES = ['./', './index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

// لا نتدخل إطلاقاً في طلبات POST (كل نداءات الخادم لـ Google Apps Script) ولا أي مصدر
// خارجي آخر — فقط نوفّر نسخة احتياطية بسيطة من واجهة التطبيق نفسها (index.html والأيقونات)
// في حال تعذّر الاتصال بالشبكة مؤقتاً، بلا أي تأثير على تحديث البيانات الفعلي.
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(req).then((res) => {
      const resClone = res.clone();
      caches.open(CACHE_NAME).then((cache) => cache.put(req, resClone)).catch(() => {});
      return res;
    }).catch(() => caches.match(req).then((cached) => cached || caches.match('./index.html')))
  );
});
