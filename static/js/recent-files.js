/* Banki UI - Recent files list (localStorage). */
(function () {
  "use strict";
  const KEY = "banki.recentFiles";
  function load() { try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch { return []; } }
  function save(list) { try { localStorage.setItem(KEY, JSON.stringify(list.slice(0, 5))); } catch {} }
  function add(name, size) {
    const list = load().filter(x => x.name !== name);
    list.unshift({ name, size, ts: Date.now() });
    save(list); render();
  }
  function render() {
    const input = document.getElementById("fileInput");
    if (!input) return;
    let host = document.getElementById("recentFilesList");
    if (!host) {
      host = document.createElement("div");
      host.id = "recentFilesList"; host.className = "recent-files";
      input.closest(".full-row, .import-grid > div")?.appendChild(host);
    }
    const items = load();
    host.innerHTML = items.length
      ? `<small>Legutóbbi fájlok:</small> ${items.map(i =>
          `<button type="button" class="chip" title="${i.name} (${Math.round(i.size/1024)} KB)" data-name="${i.name}">${i.name}</button>`).join(" ")}`
      : "";
  }
  function init() {
    const input = document.getElementById("fileInput");
    if (!input) return;
    input.addEventListener("change", () => {
      const f = input.files?.[0];
      if (f) add(f.name, f.size);
    });
    render();
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
  document.addEventListener("click", e => {
    if (e.target?.id === "openImportBtn") setTimeout(render, 60);
  });
})();
