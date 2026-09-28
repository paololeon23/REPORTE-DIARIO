const CACHE = "qb-cosecha-v135";
const ASSETS = [
  "./",
  "./index.html",
  "./install.html",
  "./css/fonts.css",
  "./css/styles.css",
  "./fonts/plus-jakarta-sans-500.woff2",
  "./fonts/plus-jakarta-sans-600.woff2",
  "./fonts/plus-jakarta-sans-700.woff2",
  "./fonts/plus-jakarta-sans-800.woff2",
  "./js/config.js",
  "./js/catalog-lotes.js",
  "./js/catalog-supervisores.js",
  "./js/catalog-plano.js",
  "./js/data.js",
  "./js/select.js",
  "./js/api.js",
  "./js/excel.js",
  "./js/app.js",
  "./manifest.json",
  "./data/supervisores-cosecha.json",
  "./data/lotes-licapa.json",
  "./data/plano-cosecha-etapa-i.json",
  "./data/trabajadores.json",
  "./assets/logo-qberries.png",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/apple-touch-icon.png",
];

function isStaticAsset(url) {
  return /\.(json|css|js|woff2|png|jpg|webp|svg)$/i.test(url.pathname);
}

function isHtmlResponse(res) {
  if (!res) return false;
  const ct = (res.headers.get("content-type") || "").toLowerCase();
  return ct.includes("text/html");
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE)
      .then((c) =>
        Promise.all(
          ASSETS.map((url) =>
            c.add(url).catch((err) => {
              console.warn("SW skip", url, err && err.message);
            })
          )
        )
      )
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (/script\.google\.com|googleusercontent\.com/i.test(url.href)) return;

  const htmlNav = req.mode === "navigate" || url.pathname.endsWith(".html") || url.pathname === "/" || url.pathname.endsWith("/");
  const code = /\.(js|css|html)$/i.test(url.pathname) || htmlNav;

  if (code) {
    event.respondWith(
      fetch(req)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => caches.match(req).then((hit) => hit || caches.match("./index.html") || caches.match("./")))
    );
    return;
  }

  const asset = isStaticAsset(url);
  event.respondWith(
    caches.match(req).then((hit) => {
      const fresh = fetch(req)
        .then((res) => {
          if (res && res.ok && asset && !isHtmlResponse(res)) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => hit || Response.error());
      return hit || fresh;
    })
  );
});
