/* Banki UI - PWA service worker registration (production only, same origin). */
(function () {
  "use strict";
  if (!("serviceWorker" in navigator)) return;
  const host = location.hostname;
  // Skip in dev/local preview contexts
  if (host === "localhost" || host === "127.0.0.1" || host.endsWith(".local")) {
    navigator.serviceWorker.getRegistrations().then(rs => rs.forEach(r => {
      if (r.scope.endsWith("/")) r.unregister();
    }));
    return;
  }
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js").catch(() => {});
  });
})();
