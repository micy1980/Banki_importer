/* Banki UI - First-run onboarding wizard.
   Shows a 3-step guide if no companies are saved. Stored dismissal in localStorage. */
(function () {
  "use strict";
  const KEY = "banki.onboarding.done";
  async function shouldShow() {
    if (localStorage.getItem(KEY)) return false;
    try {
      const r = await fetch("/api/companies");
      const data = await r.json();
      return !(data.companies && data.companies.length);
    } catch { return false; }
  }
  function build() {
    const dlg = document.createElement("dialog");
    dlg.id = "onboardDialog"; dlg.className = "preview-dialog onboard-dialog";
    let step = 0;
    const steps = [
      { title: "Üdv a Banki import konvertálóban!",
        body: "Három lépésben végigvezetünk az első importon. Ez bármikor kihagyható.",
        cta: "Kezdjük" },
      { title: "1. Cég létrehozása",
        body: "Először adj hozzá egy céget, akinek a banki importját készíted.",
        cta: "Cégek megnyitása",
        run: () => document.getElementById("companiesBtn")?.click() },
      { title: "2. Saját bankszámla rögzítése",
        body: "Vidd fel a cég bankszámláját. A 8 jegyű prefix automatikusan azonosítja a bankot az MNB táblából.",
        cta: "Bankszámlák megnyitása",
        run: () => document.getElementById("accountsBtn")?.click() },
      { title: "3. Első import",
        body: "Nyisd meg az Import panelt, válassz bankot és formátumot, majd tölts fel egy Excelt.",
        cta: "Import megnyitása",
        run: () => document.getElementById("openImportBtn")?.click() },
    ];
    function render() {
      const s = steps[step];
      dlg.innerHTML = `
        <div class="dialog-head">
          <h2>${s.title} <span class="muted-small">(${step+1}/${steps.length})</span></h2>
          <button type="button" class="secondary" data-close>Kihagyás</button>
        </div>
        <div class="dialog-body">
          <p>${s.body}</p>
          <div class="account-actions" style="justify-content:space-between;margin-top:16px;">
            <button type="button" class="secondary" data-prev ${step===0?"disabled":""}>Vissza</button>
            <button type="button" class="primary" data-next>${s.cta}</button>
          </div>
        </div>`;
      dlg.querySelector("[data-close]").onclick = finish;
      dlg.querySelector("[data-prev]").onclick = () => { step = Math.max(0, step-1); render(); };
      dlg.querySelector("[data-next]").onclick = () => {
        s.run?.();
        if (step >= steps.length-1) finish();
        else { step++; render(); }
      };
    }
    function finish() { localStorage.setItem(KEY, "1"); dlg.close(); dlg.remove(); }
    document.body.appendChild(dlg);
    render();
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open","");
  }
  async function init() {
    if (await shouldShow()) build();
    // restart trigger from help dialog
    document.addEventListener("click", e => {
      if (e.target?.dataset?.action === "restart-onboarding") {
        localStorage.removeItem(KEY); build();
      }
    });
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
