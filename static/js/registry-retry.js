/* Banki UI - MNB bank registry pill.
   Click/keyboard retry: POST /api/bank-registry/refresh.
   If online refresh fails, the API returns the saved local cache as stale data.
*/
(function () {
  "use strict";

  function fmtDate(value) {
    if (!value) return "ismeretlen id\u0151pont";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleString("hu-HU", { dateStyle: "short", timeStyle: "short" });
  }

  function rowCountOf(registry) {
    return Number(registry?.row_count || registry?.rows?.length || 0);
  }

  function describe(registry) {
    const rowCount = rowCountOf(registry);
    const hasRows = rowCount > 0;
    const stale = !!(
      registry?.error ||
      registry?.startup_ok === false ||
      registry?.status === "stale" ||
      registry?.status === "warn"
    );
    if (hasRows && !stale) return { variant: "ok", text: `MNB t\u00e1bla: ${rowCount} prefix` };
    if (hasRows) return { variant: "warn", text: "MNB t\u00e1bla: helyi cache (kattints az \u00fajrapr\u00f3b\u00e1lkoz\u00e1shoz)" };
    return { variant: "bad", text: "MNB t\u00e1bla: nem el\u00e9rhet\u0151 (kattints az \u00fajrapr\u00f3b\u00e1lkoz\u00e1shoz)" };
  }

  function setTooltip(pill, registry) {
    const parts = [];
    const rowCount = rowCountOf(registry);
    if (rowCount) parts.push(`${rowCount} prefix`);
    if (registry?.updated_at) parts.push(`friss\u00edtve: ${fmtDate(registry.updated_at)}`);
    if (registry?.error) parts.push(`hiba: ${registry.error}`);
    if (registry?.source_url) parts.push(registry.source_url);
    parts.push("kattints \u00fajrat\u00f6lt\u00e9shez");
    pill.title = parts.join(" - ");
  }

  function apply(pill, state, registry) {
    pill.classList.remove("ok", "warn", "bad", "error", "info");
    pill.classList.add(state.variant);
    pill.textContent = state.text;
    pill.setAttribute("aria-label", state.text);
    if (registry) setTooltip(pill, registry);
    if (Array.isArray(registry?.rows)) window.registryRows = registry.rows;
  }

  async function readJsonResponse(response) {
    let payload = {};
    try {
      payload = await response.json();
    } catch (_) {
      payload = {};
    }
    if (!response.ok && !rowCountOf(payload)) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload;
  }

  async function refresh(pill) {
    apply(pill, { variant: "info", text: "MNB t\u00e1bla: friss\u00edt\u00e9s..." });
    try {
      let response = await fetch("/api/bank-registry/refresh", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: "{}",
      });

      let registry = await readJsonResponse(response);
      if (!rowCountOf(registry)) {
        response = await fetch("/api/bank-registry?t=" + Date.now(), { cache: "no-store" });
        registry = await readJsonResponse(response);
      }

      const state = describe(registry);
      apply(pill, state, registry);
      window.renderRegistry?.(registry);
      window.bankiToast?.(
        state.variant === "ok"
          ? `MNB t\u00e1bla friss\u00edtve: ${rowCountOf(registry)} prefix.`
          : state.variant === "warn"
            ? "Az MNB t\u00e1bla most nem friss\u00fclt, a helyi cache haszn\u00e1latban maradt."
            : "MNB t\u00e1bla nem \u00e9rhet\u0151 el.",
        state.variant,
      );
    } catch (error) {
      apply(pill, { variant: "bad", text: "MNB t\u00e1bla: hiba (\u00fajrapr\u00f3b\u00e1l\u00e1s kattint\u00e1ssal)" });
      window.bankiToast?.("MNB lek\u00e9rdez\u00e9s sikertelen: " + (error.message || error), "error");
    }
  }

  async function initialFetch(pill) {
    try {
      const response = await fetch("/api/bank-registry?t=" + Date.now(), { cache: "no-store" });
      if (!response.ok) return;
      const registry = await response.json();
      apply(pill, describe(registry), registry);
    } catch (_) {
      /* Keep server-rendered state. */
    }
  }

  function init() {
    const pill = document.getElementById("registryPill");
    if (!pill || pill.dataset.registryRetryReady === "1") return;
    pill.dataset.registryRetryReady = "1";
    pill.setAttribute("role", "button");
    pill.setAttribute("tabindex", "0");
    pill.style.cursor = "pointer";
    pill.title = "Kattints az MNB t\u00e1bla \u00fajrat\u00f6lt\u00e9s\u00e9hez";

    pill.addEventListener("click", () => refresh(pill));
    pill.addEventListener("keydown", event => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      refresh(pill);
    });
    initialFetch(pill);
  }

  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
