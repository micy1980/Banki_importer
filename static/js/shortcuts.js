/* Banki UI - Keyboard shortcuts.
   Ctrl/Cmd+I: open Import. Ctrl/Cmd+Enter: convert. Ctrl/Cmd+K: command palette.
   Esc inside palette closes it. */
(function () {
  "use strict";
  function isTypingTarget(el) {
    if (!el) return false;
    const tag = el.tagName;
    return tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || el.isContentEditable;
  }
  function palette() {
    let dlg = document.getElementById("cmdPalette");
    if (dlg) return dlg;
    dlg = document.createElement("dialog");
    dlg.id = "cmdPalette"; dlg.className = "cmd-palette";
    dlg.innerHTML = `
      <input type="search" id="cmdPaletteInput" placeholder="Parancs vagy partner / cég kereső…" autocomplete="off">
      <ul id="cmdPaletteList" role="listbox"></ul>`;
    document.body.appendChild(dlg);
    dlg.addEventListener("click", e => { if (e.target === dlg) dlg.close(); });
    return dlg;
  }
  const COMMANDS = [
    { label: "Import megnyitása", run: () => document.getElementById("openImportBtn")?.click() },
    { label: "TXT letöltése", run: () => document.getElementById("convertBtn")?.click() },
    { label: "Összegzés", run: () => document.getElementById("summaryBtn")?.click() },
    { label: "Cégek kezelése", run: () => document.getElementById("companiesBtn")?.click() },
    { label: "Saját bankszámlák", run: () => document.getElementById("accountsBtn")?.click() },
    { label: "Partnerek", run: () => document.getElementById("partnersBtn")?.click() },
    { label: "Súgó", run: () => document.getElementById("helpBtn")?.click() },
    { label: "Téma váltása (világos/sötét)", run: () => document.getElementById("themeToggle")?.click() },
    { label: "MNB tábla újratöltése", run: () => document.getElementById("registryPill")?.click() },
  ];
  function render(query) {
    const list = document.getElementById("cmdPaletteList");
    const q = (query || "").toLowerCase().trim();
    const items = COMMANDS.filter(c => !q || c.label.toLowerCase().includes(q));
    list.innerHTML = "";
    items.forEach((c, i) => {
      const li = document.createElement("li");
      li.textContent = c.label; li.tabIndex = 0; li.dataset.idx = i;
      if (i === 0) li.classList.add("active");
      li.addEventListener("click", () => { document.getElementById("cmdPalette").close(); c.run(); });
      list.appendChild(li);
    });
  }
  function openPalette() {
    const dlg = palette();
    render("");
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open","");
    setTimeout(() => document.getElementById("cmdPaletteInput")?.focus(), 30);
  }
  document.addEventListener("keydown", e => {
    const meta = e.ctrlKey || e.metaKey;
    if (meta && (e.key === "i" || e.key === "I")) {
      e.preventDefault(); document.getElementById("openImportBtn")?.click(); return;
    }
    if (meta && e.key === "Enter") {
      e.preventDefault(); document.getElementById("convertBtn")?.click(); return;
    }
    if (meta && (e.key === "k" || e.key === "K")) {
      e.preventDefault(); openPalette(); return;
    }
  });
  document.addEventListener("input", e => {
    if (e.target?.id === "cmdPaletteInput") render(e.target.value);
  });
  document.addEventListener("keydown", e => {
    if (e.target?.id !== "cmdPaletteInput") return;
    if (e.key === "Enter") {
      const first = document.querySelector("#cmdPaletteList li.active") || document.querySelector("#cmdPaletteList li");
      first?.click();
    }
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const items = [...document.querySelectorAll("#cmdPaletteList li")];
      if (!items.length) return;
      let idx = items.findIndex(li => li.classList.contains("active"));
      items[idx]?.classList.remove("active");
      idx = (idx + (e.key === "ArrowDown" ? 1 : -1) + items.length) % items.length;
      items[idx].classList.add("active");
      items[idx].scrollIntoView({ block: "nearest" });
    }
  });
})();
