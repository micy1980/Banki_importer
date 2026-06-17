/* Banki UI - minimal static chrome i18n without page reload. */
(function () {
  "use strict";
  const KEY = "banki.lang";
  const EN = {
    "Cégek": "Companies",
    "Import": "Import",
    "Saját bankszámlák": "Own accounts",
    "Partnerek": "Partners",
    "TXT letöltése": "Download TXT",
    "Excel sablon": "Excel template",
    "? Súgó": "? Help",
    "Bezárás": "Close",
    "Mentés": "Save",
    "Mégse": "Cancel",
    "Beolvasás": "Read file",
    "Importálás": "Import",
    "Súgó": "Help",
    "Összegzés": "Summary",
    "Napló": "Log",
    "Kijelöltek törlése": "Delete selected",
    "Excel export": "Export to Excel",
    "Beolvasás eredménye": "Read result",
    "TXT előnézet": "TXT preview",
    "Bank": "Bank",
    "Formátum": "Format",
    "Fájl": "File",
    "TXT kódolás": "TXT encoding",
    "Azonosító dátuma": "Identifier date",
    "Beolvasott minta": "Sample data",
  };

  function textNodes(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        const value = node.nodeValue.trim();
        if (!value) return NodeFilter.FILTER_REJECT;
        if (node.parentElement?.closest("script,style,pre,code")) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      },
    });
    const nodes = [];
    let node;
    while ((node = walker.nextNode())) nodes.push(node);
    return nodes;
  }

  function remember(node) {
    if (node.parentElement && !node.parentElement.dataset.i18nOriginal) {
      node.parentElement.dataset.i18nOriginal = node.nodeValue;
    }
  }

  function applyLang(lang) {
    document.documentElement.lang = lang === "en" ? "en" : "hu";
    document.documentElement.dataset.lang = lang;
    for (const node of textNodes(document.body)) {
      remember(node);
      const original = node.parentElement?.dataset.i18nOriginal || node.nodeValue;
      const trimmed = original.trim();
      if (lang === "en" && EN[trimmed]) {
        node.nodeValue = original.replace(trimmed, EN[trimmed]);
      } else if (lang === "hu") {
        node.nodeValue = original;
      }
    }
    localStorage.setItem(KEY, lang);
  }

  function inject() {
    if (document.getElementById("langToggle")) return;
    const host = document.querySelector(".header-meta");
    if (!host) return;
    const button = document.createElement("button");
    button.id = "langToggle";
    button.type = "button";
    button.className = "lang-toggle";
    button.title = "Nyelv / Language";
    const lang = localStorage.getItem(KEY) || "hu";
    button.textContent = lang === "en" ? "HU" : "EN";
    button.addEventListener("click", () => {
      const next = (localStorage.getItem(KEY) || "hu") === "hu" ? "en" : "hu";
      document.documentElement.classList.add("no-ui-transition");
      applyLang(next);
      button.textContent = next === "en" ? "HU" : "EN";
      requestAnimationFrame(() => document.documentElement.classList.remove("no-ui-transition"));
    });
    host.appendChild(button);
    if (lang === "en") applyLang("en");
  }

  if (document.readyState !== "loading") inject();
  else document.addEventListener("DOMContentLoaded", inject);
})();
