/* Family FD Tracker — minimal service worker.
 *
 * Network-first for every GET: you always get the freshest app after a
 * deploy, and the last good copy is served only when the device is offline.
 * POSTs (the Apps Script API) are never intercepted or cached.
 */
var CACHE = 'fd-tracker-v1';

self.addEventListener('install', function () {
  self.skipWaiting();
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; })
        .map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return; // API traffic always hits the network
  e.respondWith(
    fetch(req).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copy); }).catch(function () {});
      return res;
    }).catch(function () {
      return caches.match(req).then(function (hit) {
        if (hit) return hit;
        throw new Error('offline and not cached');
      });
    })
  );
});
