/* Catálogo Llaza instalable: guarda la app para abrirla sin conexión (los datos los guarda app/pwa.js en «cl-datos»).
   Archivos propios: primero la red (para coger los cambios) y, si no hay, la copia. Librerías y fuentes: la copia, y se renueva detrás. */
const V = "cl-app-4";
const SHELL = ["./", "index.html", "manifest.json", "app/app.css", "app/pwa.css", "app/pwa.js", "app/client.js", "icon/icon-32.png", "icon/icon-192.png", "icon/icon-512.png"];
const CDN = /^https:\/\/(fonts\.googleapis\.com|fonts\.gstatic\.com|cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net)\//;

self.addEventListener("install", e => e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())));
self.addEventListener("activate", e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k !== V && k !== "cl-datos").map(k => caches.delete(k)))).then(() => self.clients.claim())));

self.addEventListener("fetch", e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET") return;
  if (url.origin === location.origin) {
    e.respondWith(fetch(req).then(r => { if (r.ok) { const c = r.clone(); caches.open(V).then(x => x.put(req, c)); } return r; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || (req.mode === "navigate" ? caches.match("index.html") : Response.error()))));
  } else if (CDN.test(req.url)) {
    e.respondWith(caches.open(V).then(c => c.match(req).then(hit => {
      const net = fetch(req).then(r => { if (r.ok || r.type === "opaque") c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    })));
  }
});
