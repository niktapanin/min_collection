// Service worker каталога: офлайн-оболочка + кэш данных (network-first)
const SHELL = 'shell-v2';
const DATA  = 'data-v1';
const SHELL_URLS = ['./', './index.html', './extensions.js', './manifest.json', './species_reference.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(SHELL).then(c => c.addAll(SHELL_URLS)).then(()=>self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(
    keys.filter(k => k!==SHELL && k!==DATA).map(k => caches.delete(k))
  )).then(()=>self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;

  // Supabase REST и CDN-библиотеки: network-first с кэшем на офлайн
  const isData = url.hostname.endsWith('supabase.co') && url.pathname.startsWith('/rest/');
  const isCDN  = ['cdn.tailwindcss.com','unpkg.com','d3js.org','fonts.googleapis.com','fonts.gstatic.com'].includes(url.hostname);
  if (isData || isCDN) {
    e.respondWith(
      fetch(e.request).then(res => {
        const copy = res.clone();
        caches.open(DATA).then(c => c.put(e.request, copy));
        return res;
      }).catch(() => caches.match(e.request))
    );
    return;
  }
  // Фото из Supabase Storage: cache-first (не меняются под тем же URL)
  if (url.hostname.endsWith('supabase.co') && url.pathname.includes('/storage/')) {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(res => {
      const copy = res.clone(); caches.open(DATA).then(c => c.put(e.request, copy)); return res;
    })));
    return;
  }
  // Оболочка: cache-first
  if (url.origin === location.origin) {
    e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
  }
});
