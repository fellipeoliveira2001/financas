/* Performance Financial · service worker: guarda o "esqueleto" da app para abrir depressa e mesmo sem rede.
   Os dados (pedidos ao servidor) nunca são guardados. */
const VERSAO = "pf-v3";
const FICHEIROS = ["./", "index.html", "manifest.webmanifest", "config.js", "icon-192.png", "icon-512.png", "apple-touch-icon.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(VERSAO).then(c => c.addAll(FICHEIROS)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== "GET" || u.origin !== self.location.origin) return;           // pedidos ao servidor e CDNs: sempre pela rede
  e.respondWith(fetch(r).then(res => { const c = res.clone(); caches.open(VERSAO).then(ca => ca.put(r, c)); return res; }).catch(() => caches.match(r).then(m => m || caches.match("index.html"))));
});
