self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('fetch', (event) => {
  // Service Worker mínimo apenas para ativar o critério de PWA / Add to Home Screen no Android.
});
