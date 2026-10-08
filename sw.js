const CACHE = "qb-cosecha-v210";
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
  "./data/plano-cosecha-etapa-ii.json",
  "./data/trabajadores.json",
  "./assets/logo-qberries.png",
  "./assets/logo-qberries - copia.png",
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

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((list) => {
      for (let i = 0; i < list.length; i++) {
        if (list[i].focus) return list[i].focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow("./");
    })
  );
});

function openSyncDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open("qb-cosecha-sync", 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains("job")) db.createObjectStore("job");
      if (!db.objectStoreNames.contains("done")) db.createObjectStore("done");
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error("idb"));
  });
}

function readJob() {
  return openSyncDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("job", "readonly");
    const req = tx.objectStore("job").get("current");
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error || new Error("idb"));
    tx.oncomplete = () => db.close();
  }));
}

function writeJob(job) {
  return openSyncDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("job", "readwrite");
    tx.objectStore("job").put(job, "current");
    tx.oncomplete = () => { db.close(); resolve(); };
    tx.onerror = () => { db.close(); reject(tx.error || new Error("idb")); };
  }));
}

function markDone(ids) {
  return openSyncDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction("done", "readwrite");
    const store = tx.objectStore("done");
    ids.forEach((id) => store.put({ at: Date.now() }, id));
    tx.oncomplete = () => { db.close(); resolve(); };
    tx.onerror = () => { db.close(); reject(tx.error || new Error("idb")); };
  }));
}

function postConfirmed(endpoint, records) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 50000);
  return fetch(endpoint, {
    method: "POST",
    redirect: "follow",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "batchSave", rebuild: true, records }),
    signal: ctrl.signal,
  }).then((res) => res.text().then((text) => {
    clearTimeout(timer);
    if (!res.ok || !/"ok"\s*:\s*true/.test(text)) throw new Error("bad-response");
  })).catch((err) => {
    clearTimeout(timer);
    throw err;
  });
}

async function dropSent(ids) {
  const job = await readJob();
  if (!job) return;
  const gone = {};
  ids.forEach((id) => { gone[id] = true; });
  const records = (job.records || []).filter((r) => r && !gone[r.clientId]);
  await writeJob({ endpoint: job.endpoint, records: records, at: Date.now() });
}

async function runOutboxSync() {
  let sent = 0;
  while (true) {
    const job = await readJob();
    const records = job && Array.isArray(job.records) ? job.records : [];
    if (!job || !job.endpoint || !records.length) break;
    const before = records.length;
    const batch = records.slice(0, 8);
    await postConfirmed(job.endpoint, batch);
    const ids = batch.map((r) => r && r.clientId).filter(Boolean);
    if (!ids.length) throw new Error("stuck");
    await markDone(ids);
    await dropSent(ids);
    const next = await readJob();
    const left = next && next.records ? next.records.length : 0;
    if (left >= before) throw new Error("stuck");
    sent += ids.length;
  }
  if (!sent) return;
  try {
    await self.registration.showNotification("Pendientes enviados", {
      body: "Se subieron " + sent + (sent === 1 ? " registro." : " registros."),
      icon: "./assets/icon-192.png",
      tag: "qb-envio-ok",
    });
  } catch (_) {}
  const clients = await self.clients.matchAll({ includeUncontrolled: true, type: "window" });
  clients.forEach((client) => client.postMessage({ type: "qb-sync-done" }));
}

self.addEventListener("sync", (event) => {
  if (event.tag !== "qb-cosecha-outbox") return;
  event.waitUntil(runOutboxSync());
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
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req)
          .then((res) => {
            if (res && res.ok && (htmlNav || !isHtmlResponse(res))) {
              const copy = res.clone();
              caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
            }
            return res;
          })
          .catch(() => caches.match("./index.html").then((page) => page || caches.match("./")));
      })
    );
    return;
  }

  const asset = isStaticAsset(url);
  event.respondWith(
    caches.match(req).then((hit) => {
      if (hit) return hit;
      return fetch(req)
        .then((res) => {
          if (res && res.ok && asset && !isHtmlResponse(res)) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => Response.error());
    })
  );
});
