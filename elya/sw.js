/* Esquema ELYA instalable: guarda la app para abrirla sin conexión. Es una app aparte del catálogo (caché «elya-…»).
   Archivos propios: primero la red (para coger los cambios) y, si no hay o tarda, la copia. Fuentes: la copia, y se renueva detrás. */
const V = "elya-75";
const SHELL = ["./", "index.html", "manifest.json", "icon/icon-32.png", "icon/icon-180.png", "icon/icon-192.png", "icon/icon-512.png", "lib/pdfjs/pdf.min.js", "lib/pdfjs/pdf.worker.min.js"];
const CDN = /^https:\/\/(fonts\.googleapis\.com|fonts\.gstatic\.com)\//;
/* Con wifi sin internet o mala señal la red no falla, se queda esperando: pasado este tiempo se usa la copia guardada. */
const NET_TIMEOUT = 4000;

self.addEventListener("install", e => e.waitUntil(caches.open(V).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting())));
/* Solo borra copias viejas de esta app: las del catálogo («cl-…») comparten el mismo sitio. */
self.addEventListener("activate", e => e.waitUntil(
  caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("elya-") && k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim())));

self.addEventListener("fetch", e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET") return;
  if (url.origin === location.origin) {
    /* no-cache: pregunta siempre a GitHub si hay versión nueva (si no, GitHub sirve la anterior hasta 10 min). */
    const red = fetch(req, { cache: "no-cache" }).then(r => { if (r.ok) { const c = r.clone(); caches.open(V).then(x => x.put(req, c)); } return r; });
    e.waitUntil(red.catch(() => {}));
    const copia = () => caches.match(req, { ignoreSearch: true }).then(r => r || (req.mode === "navigate" ? caches.match("index.html") : undefined));
    e.respondWith(new Promise(resolve => {
      let hecho = false;
      const dar = r => { if (!hecho && r) { hecho = true; resolve(r); } };
      const timer = setTimeout(() => copia().then(dar), NET_TIMEOUT);
      red.then(r => { clearTimeout(timer); dar(r); })
        .catch(() => { clearTimeout(timer); copia().then(c => { dar(c); if (!hecho) resolve(Response.error()); }); });
    }));
  } else if (CDN.test(req.url)) {
    e.respondWith(caches.open(V).then(c => c.match(req).then(hit => {
      const net = fetch(req).then(r => { if (r.ok || r.type === "opaque") c.put(req, r.clone()); return r; }).catch(() => hit);
      e.waitUntil(net);
      if (hit) return hit;
      /* Sin copia, no frena el arranque: sigue bajando y queda guardada para la próxima vez. */
      return Promise.race([net, new Promise(ok => setTimeout(() => ok(Response.error()), NET_TIMEOUT))]);
    })));
  }
});
