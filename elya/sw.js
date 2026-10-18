/* Esquema ELYA instalable: guarda la app para abrirla sin conexión. Es una app aparte del catálogo (caché «elya-…»).
   Archivos propios: primero la red (para coger los cambios) y, si no hay, la copia. Fuentes: la copia, y se renueva detrás. */
const V = "elya-4";
const SHELL = ["./", "index.html", "manifest.json", "icon/icon-32.png", "icon/icon-180.png", "icon/icon-192.png", "icon/icon-512.png", "lib/pdfjs/pdf.min.js", "lib/pdfjs/pdf.worker.min.js"];
const CDN = /^https:\/\/(fonts\.googleapis\.com|fonts\.gstatic\.com)\//;

self.addEventListener("install", e => e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())));
/* Solo borra copias viejas de esta app: las del catálogo («cl-…») comparten el mismo sitio. */
self.addEventListener("activate", e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("elya-") && k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())));

self.addEventListener("fetch", e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET") return;
  if (url.origin === location.origin) {
    /* no-cache: pregunta siempre a GitHub si hay versión nueva (si no, GitHub sirve la anterior hasta 10 min). */
    e.respondWith(fetch(req, { cache: "no-cache" }).then(r => { if (r.ok) { const c = r.clone(); caches.open(V).then(x => x.put(req, c)); } return r; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || (req.mode === "navigate" ? caches.match("index.html") : Response.error()))));
  } else if (CDN.test(req.url)) {
    e.respondWith(caches.open(V).then(c => c.match(req).then(hit => {
      const net = fetch(req).then(r => { if (r.ok || r.type === "opaque") c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    })));
  }
});
