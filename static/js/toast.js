/* Banki UI - Toast notifications. Exposes window.bankiToast(msg, variant, ms). */
(function () {
  "use strict";
  function host() {
    let h = document.getElementById("toastHost");
    if (!h) {
      h = document.createElement("div");
      h.id = "toastHost";
      h.setAttribute("role", "region");
      h.setAttribute("aria-live", "polite");
      h.setAttribute("aria-label", "Értesítések");
      document.body.appendChild(h);
    }
    return h;
  }
  function toast(message, variant = "info", timeoutMs = 3800) {
    const n = document.createElement("div");
    n.className = `toast toast--${variant}`;
    n.setAttribute("role", variant === "error" ? "alert" : "status");
    n.textContent = message;
    host().appendChild(n);
    requestAnimationFrame(() => n.classList.add("toast--in"));
    const close = () => {
      n.classList.remove("toast--in");
      n.addEventListener("transitionend", () => n.remove(), { once: true });
      setTimeout(() => n.remove(), 400);
    };
    setTimeout(close, timeoutMs);
    n.addEventListener("click", close);
    return close;
  }
  window.bankiToast = toast;
})();
