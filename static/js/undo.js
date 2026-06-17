/* Banki UI - Undo toast. Extends toast with action button + handler. */
(function () {
  "use strict";
  function ensureHost() {
    let h = document.getElementById("toastHost");
    if (!h) {
      h = document.createElement("div");
      h.id = "toastHost"; h.setAttribute("role","region");
      h.setAttribute("aria-live","polite"); h.setAttribute("aria-label","Értesítések");
      document.body.appendChild(h);
    }
    return h;
  }
  function undoToast(message, action, timeoutMs = 6000) {
    const n = document.createElement("div");
    n.className = "toast toast--info toast--with-action";
    n.setAttribute("role","status");
    const msg = document.createElement("span"); msg.textContent = message;
    const btn = document.createElement("button");
    btn.type = "button"; btn.className = "toast-action"; btn.textContent = "Visszavonás";
    n.append(msg, btn);
    ensureHost().appendChild(n);
    requestAnimationFrame(()=>n.classList.add("toast--in"));
    let done = false;
    const close = () => {
      if (done) return; done = true;
      n.classList.remove("toast--in");
      setTimeout(()=>n.remove(), 400);
    };
    btn.addEventListener("click", async () => {
      btn.disabled = true; btn.textContent = "Visszaállítás…";
      try { await action(); window.bankiToast?.("Visszaállítva.", "ok"); }
      catch (e) { window.bankiToast?.("Visszaállítás nem sikerült: " + (e.message || e), "error"); }
      close();
    });
    setTimeout(close, timeoutMs);
    return close;
  }
  window.bankiUndoToast = undoToast;

  // Wrap delete buttons: intercept /api/*/delete network responses via fetch wrapper
  // to capture the deleted entity for restore. Minimal, safe.
  const origFetch = window.fetch.bind(window);
  window.fetch = async function (input, init) {
    const url = typeof input === "string" ? input : (input?.url || "");
    const isDelete = /\/api\/(partners|accounts|companies)\/delete/.test(url);
    let body = null;
    if (isDelete && init?.body) {
      try { body = JSON.parse(init.body); } catch (_) {}
    }
    const res = await origFetch(input, init);
    if (isDelete && body && res.ok) {
      const kind = url.match(/\/api\/(\w+)\/delete/)[1];
      const labels = { partners: "Partner", accounts: "Bankszámla", companies: "Cég" };
      undoToast(`${labels[kind]} törölve.`, async () => {
        const r = await origFetch(`/api/${kind}/restore`, {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ id: body.id, company_id: body.company_id }),
        });
        if (!r.ok) throw new Error("HTTP " + r.status);
        // refresh: click the button that opens the dialog to force re-render
        document.getElementById({partners:"partnersBtn",accounts:"accountsBtn",companies:"companiesBtn"}[kind])?.click();
      });
    }
    return res;
  };
})();
