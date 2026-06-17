/* Banki UI - persisted light/dark theme toggle without visible re-render flash. */
(function () {
  "use strict";
  const KEY = "banki.theme";

  function setTheme(theme) {
    const dark = theme === "dark";
    document.documentElement.classList.toggle("theme-dark", dark);
    localStorage.setItem(KEY, dark ? "dark" : "light");
  }

  function lockTransitions() {
    document.documentElement.classList.add("no-ui-transition");
    window.setTimeout(() => document.documentElement.classList.remove("no-ui-transition"), 90);
  }

  document.addEventListener("DOMContentLoaded", () => {
    const meta = document.querySelector(".shell > header .header-meta");
    if (!meta || meta.querySelector("#themeToggle")) return;
    const button = document.createElement("button");
    button.id = "themeToggle";
    button.type = "button";
    button.className = "ghost icon";
    button.title = "Téma váltása";
    button.setAttribute("aria-label", "Téma váltása");
    button.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>`;
    button.addEventListener("click", () => {
      lockTransitions();
      setTheme(document.documentElement.classList.contains("theme-dark") ? "light" : "dark");
    });
    meta.appendChild(button);
  });
})();
