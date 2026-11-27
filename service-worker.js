'use strict';
const CACHE_NAME = 'brivatedocs-shell-v2.2.0';
const APP_SHELL = [
  './',
  './index.html',
  './style.css',
  './style-v2.css',
  './brivatedocs-logo.png',
  './brivatedocs-icon-192.png',
  './brivatedocs-icon-512.png',
  './manifest.webmanifest',
  './jszip.min.js',
  './pdf.min.js',
  './pdf-lib.min.js',
  './pdf.worker.min.js',
  './js/core.js',
  './js/ooxml-templates.js',
  './js/pptx-reader.js',
  './js/pptx-ops.js',
  './js/pptx-editor.js',
  './js/pptx-slideshow.js',
  './js/docx-editor.js',
  './js/app-shell.js',
  './js/pdf-editor.js',
  './js/image-to-pdf.js',
  './js/pdf-split.js',
  './js/pdf-tools.js',
  './js/phase1-features.js',
  './js/phase2-tools.js',
  './qpdf.js',
  './qpdf.wasm',
  './js/pwa.js',
  './js/v2-ui.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

// Network-first for HTML/CSS/JS so updates appear immediately online while remaining 100% offline-capable.
self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET' || new URL(request.url).origin !== self.location.origin) return;
  event.respondWith(
    fetch(request)
      .then((response) => {
        if (response && response.ok && response.type === 'basic') {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        }
        return response;
      })
      .catch(() =>
        caches.match(request).then((cached) => {
          if (cached) return cached;
          return request.mode === 'navigate' ? caches.match('./index.html') : Response.error();
        })
      )
  );
});
