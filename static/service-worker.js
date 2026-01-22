const CACHE_NAME = "sevens-pwa-v3";

const ASSETS = [
  "/",                 // يجيب Login.html
  "/Login.html",
  "/Portal.html",

  // صفحات مهمة حسب نظامك
  "/EmployeePage.html",
  "/DepartmentManagerPage.html",
  "/GeneralManager.html",
  "/HrPage.html",
  "/admin.html",

  // PM
  "/pm/work",
  "/pm/dashboard",

  // ملفات ثابتة
  "/static/manifest.json",
  "/static/7s.jpg",
  "/static/chatbot.png"
];

// install
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// activate
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.map((k) => (k !== CACHE_NAME ? caches.delete(k) : null)))
    )
  );
  self.clients.claim();
});

// fetch
self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // لا تكاش API ولا POST
  if (url.pathname.startsWith("/api/") || req.method !== "GET") {
    return;
  }

  // cache-first للملفات الثابتة
  if (url.pathname.startsWith("/static/") || url.pathname.startsWith("/uploads/") || url.pathname.startsWith("/chat_uploads/")) {
    event.respondWith(
      caches.match(req).then((cached) => cached || fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
        return res;
      }))
    );
    return;
  }

  // network-first للصفحات (عشان تظهر آخر تحديثات)
  event.respondWith(
    fetch(req).then((res) => {
      const copy = res.clone();
      caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
      return res;
    }).catch(() => caches.match(req).then((cached) => cached || caches.match("/Login.html")))
  );
});


self.addEventListener("message", (event) => {
  if (event.data === "skipWaiting") {
    self.skipWaiting();
  }
});
