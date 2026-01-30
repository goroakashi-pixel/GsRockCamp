const CACHE_NAME = "gsrockcamp-v1";

// 最低限だけ“事前キャッシュ”して、残り（mp3等）はアクセスされたらキャッシュ
const PRECACHE = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./service-worker.js",
  "./assets/img/nene.png",
  "./assets/icons/icon-192.png",
  "./assets/icons/icon-512.png",
  "./assets/icons/apple-touch-icon.png",
  "./games/mode/index.html",
  "./games/qualities/index.html",
  "./games/diatonic/index.html"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(PRECACHE))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.map((k) => (k === CACHE_NAME ? null : caches.delete(k))))
    )
  );
  self.clients.claim();
});

// Cache-first（オフライン強い）＋ 取れたら保存
self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // 同一オリジンだけ対象
  if (url.origin !== location.origin) return;

  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;

      return fetch(req).then((res) => {
        // mp3/png/css/js/html などは保存してオフライン化
        const type = res.headers.get("content-type") || "";
        const shouldCache =
          req.method === "GET" &&
          (type.includes("text/") ||
           type.includes("javascript") ||
           type.includes("json") ||
           type.includes("image/") ||
           type.includes("audio/"));

        if (shouldCache) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        }
        return res;
      }).catch(() => {
        // ナビゲーションが落ちたらポータルへ
        if (req.mode === "navigate") return caches.match("./index.html");
        throw new Error("Offline and not cached");
      });
    })
  );
});
