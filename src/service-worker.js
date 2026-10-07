const CACHE_NAME = "medicion-productos-v3";

const URLS_TO_CACHE = ["/", "/index.html", "/manifest.json"];

// ================================
// INSTALACIÓN
// ================================
self.addEventListener("install", (event) => {
  console.log("Service Worker: instalando versión:", CACHE_NAME);

  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(URLS_TO_CACHE);
    })
  );

  // Activar inmediatamente
  self.skipWaiting();
});

// ================================
// ACTIVACIÓN
// ================================
self.addEventListener("activate", (event) => {
  console.log("Service Worker: activando versión:", CACHE_NAME);

  event.waitUntil(
    caches
      .keys()
      .then((cacheNames) => {
        return Promise.all(
          cacheNames
            .filter((cacheName) => cacheName !== CACHE_NAME)
            .map((cacheName) => {
              console.log("Eliminando caché antigua:", cacheName);

              return caches.delete(cacheName);
            })
        );
      })
      .then(() => {
        return self.clients.claim();
      })
  );
});

// ================================
// PETICIONES
// ================================
self.addEventListener("fetch", (event) => {
  // Solo GET
  if (event.request.method !== "GET") {
    return;
  }

  // NO interceptar Supabase
  if (
    event.request.url.includes("supabase.co") ||
    event.request.url.includes("/rest/v1/")
  ) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Guardar solamente respuestas válidas
        if (response && response.status === 200 && response.type === "basic") {
          const responseClone = response.clone();

          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseClone);
          });
        }

        return response;
      })
      .catch(() => {
        // Si no hay internet,
        // intentar usar la caché
        return caches.match(event.request);
      })
  );
});
