/* Keeps the app's own files on the device so it opens fast and offline.
   Team data from SharePoint is never cached here. */
var VERSION = 'mdy-team-2026-10-20-01';
var SHELL = ['./', 'index.html', 'app.css', 'app.js', 'config.js', 'manifest.webmanifest',
  'vendor/msal-browser.min.js', 'vendor/jszip.min.js', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/favicon.png', 'icons/chaim-levy.jpg'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // Graph, sign-in, fonts: straight to the network
  // Network first, so updates arrive as soon as they are published; the cache covers offline.
  e.respondWith(fetch(req).then(function (res) {
    if (res.ok) { var copy = res.clone(); caches.open(VERSION).then(function (c) { c.put(req, copy); }); }
    return res;
  }).catch(function () {
    return caches.match(req).then(function (hit) { return hit || caches.match('index.html'); });
  }));
});
