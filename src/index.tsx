import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles.css";

const root = ReactDOM.createRoot(
  document.getElementById("root") as HTMLElement
);

root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// =====================================================
// SERVICE WORKER
// =====================================================

if ("serviceWorker" in navigator) {
  window.addEventListener("load", async () => {
    try {
      /*
       * =================================================
       * LIMPIEZA ÚNICA DE VERSIONES ANTIGUAS
       * =================================================
       *
       * Esto permite eliminar la versión antigua
       * que pueda estar instalada en el celular.
       */

      const RESET_KEY = "medicion-productos-reset-v3";

      const yaSeHizoLimpieza = localStorage.getItem(RESET_KEY);

      if (!yaSeHizoLimpieza) {
        console.log("Limpiando versiones antiguas...");

        // Obtener todos los Service Workers
        const registrations = await navigator.serviceWorker.getRegistrations();

        // Desregistrar versiones anteriores
        for (const registration of registrations) {
          try {
            await registration.unregister();

            console.log("Service Worker anterior eliminado.");
          } catch (error) {
            console.error("No se pudo eliminar Service Worker:", error);
          }
        }

        // Obtener todas las cachés
        const cacheNames = await caches.keys();

        // Eliminar cachés antiguas
        await Promise.all(
          cacheNames.map(async (cacheName) => {
            console.log("Eliminando caché:", cacheName);

            await caches.delete(cacheName);
          })
        );

        // Marcar la limpieza como realizada
        localStorage.setItem(RESET_KEY, "1");

        console.log("Limpieza terminada.");

        /*
         * Recargar la aplicación.
         *
         * En la siguiente carga se registrará
         * nuevamente el Service Worker nuevo.
         */
        window.location.reload();

        return;
      }

      // =================================================
      // REGISTRAR SERVICE WORKER NUEVO
      // =================================================

      const registration = await navigator.serviceWorker.register(
        "/service-worker.js?v=3",
        {
          updateViaCache: "none",
        }
      );

      console.log("Service Worker registrado:", registration.scope);

      // =================================================
      // FORZAR ACTUALIZACIÓN
      // =================================================

      await registration.update();

      console.log("Service Worker actualizado correctamente.");

      // =================================================
      // DETECTAR NUEVA VERSIÓN
      // =================================================

      registration.addEventListener("updatefound", () => {
        const newWorker = registration.installing;

        if (!newWorker) {
          return;
        }

        newWorker.addEventListener("statechange", () => {
          if (newWorker.state === "installed") {
            if (navigator.serviceWorker.controller) {
              console.log("Nueva versión disponible. Recargando...");

              window.location.reload();
            }
          }
        });
      });

      // =================================================
      // CUANDO EL SERVICE WORKER CAMBIA
      // =================================================

      let refreshing = false;

      navigator.serviceWorker.addEventListener("controllerchange", () => {
        if (refreshing) {
          return;
        }

        refreshing = true;

        console.log("Nuevo Service Worker activo. Recargando...");

        window.location.reload();
      });
    } catch (error) {
      console.error("Error con el Service Worker:", error);
    }
  });
}
