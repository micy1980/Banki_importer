/* Banki UI - MNB bank registry pill.
   - Click: POST /api/bank-registry/refresh (real re-fetch from MNB source).
   - Fallback: GET /api/bank-registry (cache) if refresh endpoint is unavailable.
   - Tooltip: shows last updated_at + row_count.
*/
(function () {
  "use strict";
  function fmtDate(iso) {
    if (!iso) return "ismeretlen időpont";
    const d = new Date(iso);
    if (isNaN(d)) return iso;
    return d.toLocaleString("hu-HU", { dateStyle: "short", timeStyle: "short" });
  }
  function describe(reg) {
    const ok = reg && reg.rows_loaded === true && (reg.row_count || 0) > 0;
    const stale = reg && (reg.error || reg.status === "stale" || reg.status === "warn");
    if (ok && !stale) return { variant: "ok",   text: `MNB tábla: ${reg.row_count} bank` };
    if (ok &&  stale) return { variant: "warn", text: "MNB tábla: helyi gyorsítótár (kattints az újrapróbálkozáshoz)" };
    return { variant: "bad", text: "MNB tábla: nem elérhető (kattints az újrapróbálkozáshoz)" };
  }
  function setTooltip(pill, reg) {
    const parts = [];
    if (reg && reg.row_count != null) parts.push(`${reg.row_count} prefix`);
    if (reg && reg.updated_at) parts.push(`frissítve: ${fmtDate(reg.updated_at)}`);
    if (reg && reg.source_url) parts.push(reg.source_url);
    parts.push("kattints újratöltéshez");
    pill.title = parts.join(" · ");
  }
  function apply(pill, state, reg) {
    pill.classList.remove("ok", "warn", "bad", "error", "info");
    pill.classList.add(state.variant);
    pill.textContent = state.text;
    if (reg) setTooltip(pill, reg);
  }
  async function refresh(pill) {
    apply(pill, { variant: "info", text: "MNB tábla: frissítés…" });
    let reg = null;
    try {
      // Real re-fetch from MNB source
      let r = await fetch("/api/bank-registry/refresh", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
      if (!r.ok) {
        // Fallback: just re-read cache
        r = await fetch("/api/bank-registry?t=" + Date.now(), { cache: "no-store" });
      }
      if (!r.ok) throw new Error("HTTP " + r.status);
      reg = await r.json();
      const st = describe(reg);
      apply(pill, st, reg);
      window.bankiToast?.(
        st.variant === "ok"   ? `MNB tábla frissítve: ${reg.row_count} bank.` :
        st.variant === "warn" ? "Helyi MNB cache használatban." :
                                "MNB tábla nem érhető el.",
        st.variant
      );
    } catch (e) {
      apply(pill, { variant: "bad", text: "MNB tábla: hiba (újrapróbálkozás kattintással)" });
      window.bankiToast?.("MNB lekérdezés sikertelen: " + (e.message || e), "error");
    }
  }
  async function initialFetch(pill) {
    try {
      const r = await fetch("/api/bank-registry", { cache: "no-store" });
      if (r.ok) {
        const reg = await r.json();
        apply(pill, describe(reg), reg);
      }
    } catch (_) { /* leave default state */ }
  }
  function init() {
    const pill = document.getElementById("registryPill");
    if (!pill) return;
    pill.setAttribute("role", "button");
    pill.setAttribute("tabindex", "0");
    pill.style.cursor = "pointer";
    pill.title = "Kattints az MNB tábla újratöltéséhez";
    const handler = () => refresh(pill);
    pill.addEventListener("click", handler);
    pill.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handler(); } });
    initialFetch(pill);
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
