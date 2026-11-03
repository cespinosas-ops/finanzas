// Guarda la app en el teléfono para que abra aunque no haya señal.
const CACHE = "finanzas-v2";
const SHELL = ["./", "index.html", "app.js", "lectores.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "apple-touch-icon.png"];
const CDN = [
  "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.js",
  "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL).then(() => Promise.all(CDN.map((u) => c.add(u).catch(() => {}))))));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.hostname.endsWith("supabase.co")) return; // datos siempre en vivo
  const propio = url.origin === location.origin;
  if (req.mode === "navigate" || (propio && /\.(js|html)$/.test(url.pathname))) {
    // la app: primero internet (para tener la última versión), si no hay, la guardada
    e.respondWith(fetch(req).then((r) => { if (r.ok) { const c = r.clone(); caches.open(CACHE).then((x) => x.put(req.mode === "navigate" ? "index.html" : req, c)); } return r; })
      .catch(() => caches.match(req.mode === "navigate" ? "index.html" : req, { ignoreSearch: true })));
    return;
  }
  e.respondWith(caches.match(req).then((hit) => hit || fetch(req).then((r) => {
    if (r.ok && (propio || /cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.g/.test(url.hostname))) {
      const c = r.clone(); caches.open(CACHE).then((x) => x.put(req, c));
    }
    return r;
  })));
});
