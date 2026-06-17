/* Banki UI - TXT diff viewer.
   Compares the current TXT preview against the previous successful export
   stored in localStorage by export-summary.js. Adds a "Diff" button to the
   preview toolbar when a previous export exists.
*/
(function () {
  "use strict";

  function diffLines(a, b) {
    // Simple line-based LCS diff. O(n*m) - fine for typical PAYORD sizes
    // (a few hundred lines). For very large files we fall back to a hash set.
    const al = a.split(/\r\n|\r|\n/);
    const bl = b.split(/\r\n|\r|\n/);
    if (al.length * bl.length > 250000) return setDiff(al, bl);
    const n = al.length, m = bl.length;
    const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
    for (let i = n - 1; i >= 0; i--)
      for (let j = m - 1; j >= 0; j--)
        dp[i][j] = al[i] === bl[j] ? dp[i + 1][j + 1] + 1
                                   : Math.max(dp[i + 1][j], dp[i][j + 1]);
    const out = [];
    let i = 0, j = 0;
    while (i < n && j < m) {
      if (al[i] === bl[j]) { out.push({ k: "=", s: al[i] }); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) { out.push({ k: "-", s: al[i] }); i++; }
      else { out.push({ k: "+", s: bl[j] }); j++; }
    }
    while (i < n) out.push({ k: "-", s: al[i++] });
    while (j < m) out.push({ k: "+", s: bl[j++] });
    return out;
  }
  function setDiff(al, bl) {
    const setA = new Set(al), setB = new Set(bl);
    const out = [];
    for (const l of al) if (!setB.has(l)) out.push({ k: "-", s: l });
    for (const l of bl) if (!setA.has(l)) out.push({ k: "+", s: l });
    return out;
  }

  function ensureDialog() {
    let dlg = document.getElementById("diffDialog");
    if (dlg) return dlg;
    dlg = document.createElement("dialog");
    dlg.id = "diffDialog";
    dlg.className = "preview-dialog diff-dialog";
    dlg.innerHTML = `
      <div class="dialog-head">
        <h2>TXT diff (előző export ↔ jelenlegi)</h2>
        <button type="button" class="secondary" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <div class="diff-legend">
          <span class="diff-add">+ új sor</span>
          <span class="diff-del">- eltűnt sor</span>
          <span class="diff-eq">= változatlan</span>
          <span id="diffStats" class="muted-small" style="margin-left:auto;"></span>
        </div>
        <pre id="diffBody" class="diff-body"></pre>
      </div>`;
    document.body.appendChild(dlg);
    dlg.addEventListener("click", e => {
      if (e.target === dlg || e.target.dataset.close !== undefined) dlg.close();
    });
    dlg.addEventListener("cancel", () => dlg.close());
    return dlg;
  }

  function show() {
    const prev = localStorage.getItem("banki.prevExport.text");
    const cur = localStorage.getItem("banki.lastExport.text");
    if (!prev || !cur) {
      window.bankiToast?.("Még nincs két export az összehasonlításhoz.", "warn");
      return;
    }
    const ops = diffLines(prev, cur);
    const added = ops.filter(o => o.k === "+").length;
    const removed = ops.filter(o => o.k === "-").length;
    const eq = ops.filter(o => o.k === "=").length;
    const dlg = ensureDialog();
    const body = dlg.querySelector("#diffBody");
    body.innerHTML = "";
    const max = 2000;
    const shown = ops.slice(0, max);
    body.append(...shown.map(o => {
      const span = document.createElement("span");
      span.className = o.k === "+" ? "diff-add" : o.k === "-" ? "diff-del" : "diff-eq";
      span.textContent = (o.k === "=" ? "  " : o.k + " ") + (o.s || "") + "\n";
      return span;
    }));
    if (ops.length > max) {
      const more = document.createElement("span");
      more.className = "muted-small";
      more.textContent = `\n… további ${ops.length - max} sor elrejtve.`;
      body.appendChild(more);
    }
    dlg.querySelector("#diffStats").textContent =
      `+${added} új · -${removed} eltűnt · ${eq} változatlan`;
    if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
  }

  function injectButton() {
    const tools = document.querySelector(".preview-tools");
    if (!tools || document.getElementById("diffBtn")) return;
    const prev = localStorage.getItem("banki.prevExport.text");
    const cur = localStorage.getItem("banki.lastExport.text");
    if (!prev || !cur) return; // only show when there's something to diff
    const btn = document.createElement("button");
    btn.type = "button"; btn.id = "diffBtn"; btn.className = "ghost";
    btn.title = "Előző exporthoz képest mi változott?";
    btn.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
      Diff`;
    btn.addEventListener("click", show);
    tools.appendChild(btn);
  }

  function init() {
    // Try a few times after load to catch the preview tools injection from preview.js
    injectButton();
    const obs = new MutationObserver(injectButton);
    obs.observe(document.body, { childList: true, subtree: true });
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
