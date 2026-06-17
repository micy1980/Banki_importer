/* Banki UI - Theme toggle (light/dark), persisted in localStorage. */
(function () {
  "use strict";
  const KEY = "banki.theme";
  if (localStorage.getItem(KEY) === "dark") document.documentElement.classList.add("theme-dark");
  document.addEventListener("DOMContentLoaded", () => {
    const meta = document.querySelector(".shell > header .header-meta");
    if (!meta || meta.querySelector("#themeToggle")) return;
    const btn = document.createElement("button");
    btn.id = "themeToggle";
    btn.type = "button";
    btn.className = "ghost icon";
    btn.title = "Téma váltása";
    btn.setAttribute("aria-label", "Téma váltása");
    btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>`;
    btn.addEventListener("click", () => {
      const dark = document.documentElement.classList.toggle("theme-dark");
      localStorage.setItem(KEY, dark ? "dark" : "light");
    });
    meta.appendChild(btn);
  });
})();
