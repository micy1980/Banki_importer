/* Banki UI - Minimal i18n. HU is the source; EN is an overlay applied at toggle.
   Only translates static UI chrome (buttons, headers); banking field labels stay HU. */
(function () {
  "use strict";
  const KEY = "banki.lang";
  const EN = {
    "Cégek": "Companies", "Import": "Import",
    "Saját bankszámlák": "Own accounts", "Partnerek": "Partners",
    "TXT letöltése": "Download TXT", "Excel sablon": "Excel template",
    "? Súgó": "? Help", "Bezárás": "Close", "Mentés": "Save",
    "Mégse": "Cancel", "Beolvasás": "Read file", "Importálás": "Import",
    "Súgó": "Help", "Összegzés": "Summary", "Visszavonás": "Undo",
    "Kijelöltek törlése": "Delete selected", "Excel export": "Export to Excel",
    "Beolvasás eredménye": "Read result", "TXT előnézet": "TXT preview",
    "Kezdjük": "Start", "Kihagyás": "Skip", "Vissza": "Back",
    "Bank": "Bank", "Formátum": "Format", "Fájl": "File",
    "TXT kódolás": "TXT encoding", "Azonosító dátuma": "Identifier date",
    "Beolvasott minta": "Sample data",
  };
  function applyLang(lang) {
    if (lang === "hu") { location.reload(); return; }
    // Walk text nodes; substitute exact matches in the HU dictionary
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let n; const subs = [];
    while ((n = walker.nextNode())) {
      const t = n.nodeValue.trim();
      if (t && EN[t]) subs.push([n, EN[t]]);
    }
    subs.forEach(([n, v]) => n.nodeValue = n.nodeValue.replace(n.nodeValue.trim(), v));
    document.documentElement.lang = "en";
    localStorage.setItem(KEY, "en");
  }
  function inject() {
    if (document.getElementById("langToggle")) return;
    const host = document.querySelector(".header-meta");
    if (!host) return;
    const btn = document.createElement("button");
    btn.id = "langToggle"; btn.type = "button"; btn.className = "lang-toggle";
    btn.title = "Nyelv / Language";
    btn.textContent = (localStorage.getItem(KEY) === "en") ? "HU" : "EN";
    btn.addEventListener("click", () => {
      const cur = localStorage.getItem(KEY) || "hu";
      const next = cur === "hu" ? "en" : "hu";
      localStorage.setItem(KEY, next);
      applyLang(next);
      btn.textContent = next === "en" ? "HU" : "EN";
    });
    host.appendChild(btn);
    if (localStorage.getItem(KEY) === "en") applyLang("en");
  }
  if (document.readyState !== "loading") inject();
  else document.addEventListener("DOMContentLoaded", inject);
})();
