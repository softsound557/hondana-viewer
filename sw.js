const SHELL_CACHE = "shell-v2";
const DATA_CACHE = "data-v2";
const SHELL_FILES = [
  "index.html",
  "app.js",
  "style.css",
  "manifest.json",
  "icons/icon-192.png",
  "icons/icon-512.png",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2/lib/msal-browser.min.js",
];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(SHELL_CACHE).then((cache) => cache.addAll(SHELL_FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== SHELL_CACHE && k !== DATA_CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.mode === "navigate" && new URL(event.request.url).origin !== self.location.origin) {
    return; // サインインのリダイレクトなど、他オリジンへの画面遷移はそのまま素通しする
  }

  const url = new URL(event.request.url);

  if (url.origin === "https://graph.microsoft.com") {
    // データ・画像はネットワーク優先。取れなければ前回取得分を返す（オフライン閲覧用）
    event.respondWith(
      fetch(event.request)
        .then((res) => {
          const clone = res.clone();
          caches.open(DATA_CACHE).then((cache) => cache.put(event.request, clone));
          return res;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // アプリ本体（画面・スクリプト）はキャッシュ優先
  event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request)));
});
