/* Service worker: saves app files on the device so the app opens without internet */
const CACHE_NAME = "qc-status-v1";     // change v1 -> v2 if you ever want to force a full refresh

// Files saved on first visit (a missing file will not break the install)
const APP_FILES = [
    "./",
    "index.html",
    "index.css",
    "index.js",
    "manifest.webmanifest",
    "qc-1.png",
    "qc-2.png",
    "icon-192.png",
    "icon-512.png"
];

// Libraries loaded from the internet (saved too, so they work offline)
const CDN_FILES = [
    "https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css",
    "https://code.jquery.com/jquery-3.5.1.slim.min.js",
    "https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js",
    "https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js",
    "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
];

const CDN_HOSTS = [
    "stackpath.bootstrapcdn.com", "code.jquery.com", "cdn.jsdelivr.net",
    "cdnjs.cloudflare.com", "fonts.googleapis.com", "fonts.gstatic.com"
];

self.addEventListener("install", function (event) {
    event.waitUntil(
        caches.open(CACHE_NAME).then(function (cache) {
            return Promise.all(
                APP_FILES.concat(CDN_FILES).map(function (url) {
                    return cache.add(url).catch(function () { /* skip files that fail */ });
                })
            );
        }).then(function () { return self.skipWaiting(); })
    );
});

self.addEventListener("activate", function (event) {
    event.waitUntil(
        caches.keys().then(function (names) {
            return Promise.all(
                names.filter(function (n) { return n !== CACHE_NAME; })
                     .map(function (n) { return caches.delete(n); })
            );
        }).then(function () { return self.clients.claim(); })
    );
});

function saveCopy(request, response) {
    if (response && (response.ok || response.type === "opaque")) {
        const copy = response.clone();
        caches.open(CACHE_NAME).then(function (cache) { cache.put(request, copy); });
    }
    return response;
}

self.addEventListener("fetch", function (event) {
    const request = event.request;
    if (request.method !== "GET") return;

    const url = new URL(request.url);

    // Google Sheet data is NOT handled here (index.js keeps its own offline copy)
    if (url.hostname.indexOf("docs.google.com") > -1 || url.hostname.indexOf("googleusercontent.com") > -1) return;

    // Libraries and fonts: saved copy first (they never change)
    if (CDN_HOSTS.indexOf(url.hostname) > -1) {
        event.respondWith(
            caches.match(request).then(function (cached) {
                return cached || fetch(request).then(function (res) { return saveCopy(request, res); });
            })
        );
        return;
    }

    // Your own files: internet first (so updates show up), saved copy when offline
    if (url.origin === self.location.origin) {
        event.respondWith(
            fetch(request).then(function (res) { return saveCopy(request, res); })
                .catch(function () {
                    return caches.match(request).then(function (cached) {
                        return cached || caches.match("index.html");
                    });
                })
        );
    }
});
