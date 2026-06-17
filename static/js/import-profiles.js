/* Banki UI - Saved import profiles (localStorage).
   A profile = { name, bank, format, encoding, mapping, defaults }. */
(function () {
  "use strict";
  const KEY = "banki.importProfiles";
  function load() { try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch { return []; } }
  function save(list) { try { localStorage.setItem(KEY, JSON.stringify(list)); } catch {} }
  function currentConfig() {
    if (typeof window.collectConfig === "function") return window.collectConfig();
    return null;
  }
  function applyConfig(p) {
    const set = (id, v) => { const el = document.getElementById(id); if (el && v != null) { el.value = v; el.dispatchEvent(new Event("change", {bubbles:true})); } };
    set("bankSelect", p.bank);
    setTimeout(() => set("formatSelect", p.format), 50);
    set("encoding", p.encoding);
    // mapping/defaults restoration happens after format selection rebuilds inputs
    setTimeout(() => {
      Object.entries(p.mapping || {}).forEach(([k,v]) => {
        const sel = document.querySelector(`[data-map="${k}"]`); if (sel) sel.value = v;
      });
      Object.entries(p.defaults || {}).forEach(([k,v]) => {
        const inp = document.querySelector(`[data-default="${k}"]`); if (inp) inp.value = v;
      });
    }, 200);
  }
  function render() {
    const host = document.getElementById("importProfilesHost");
    if (!host) return;
    const items = load();
    host.innerHTML = `
      <div class="profile-row">
        <select id="profileSelect"><option value="">Profil választása…</option>${items.map((p,i)=>`<option value="${i}">${p.name}</option>`).join("")}</select>
        <button type="button" class="secondary" id="loadProfileBtn">Betöltés</button>
        <button type="button" class="secondary" id="saveProfileBtn">Mentés profilként</button>
        <button type="button" class="ghost" id="deleteProfileBtn">Törlés</button>
      </div>`;
    document.getElementById("loadProfileBtn").onclick = () => {
      const idx = +document.getElementById("profileSelect").value;
      const p = load()[idx]; if (!p) return;
      applyConfig(p); window.bankiToast?.(`Profil betöltve: ${p.name}`, "ok");
    };
    document.getElementById("saveProfileBtn").onclick = () => {
      const name = prompt("Profil neve:");
      if (!name) return;
      const cfg = currentConfig(); if (!cfg) return;
      const list = load().filter(p => p.name !== name);
      list.push({ name, ...cfg }); save(list); render();
      window.bankiToast?.(`Profil mentve: ${name}`, "ok");
    };
    document.getElementById("deleteProfileBtn").onclick = () => {
      const idx = +document.getElementById("profileSelect").value;
      const list = load(); if (!list[idx]) return;
      if (!confirm(`"${list[idx].name}" törlése?`)) return;
      list.splice(idx, 1); save(list); render();
    };
  }
  function inject() {
    const grid = document.querySelector("#importDialog .import-grid");
    if (!grid || document.getElementById("importProfilesHost")) return;
    const host = document.createElement("div");
    host.id = "importProfilesHost"; host.className = "full-row import-profiles";
    grid.appendChild(host); render();
  }
  document.addEventListener("click", e => {
    if (e.target?.id === "openImportBtn") setTimeout(inject, 80);
  });
})();
