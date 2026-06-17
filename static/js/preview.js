/* Banki UI - TXT preview: copy-to-clipboard + full-view dialog. */
(function () {
  "use strict";
  function init() {
    const pre = document.getElementById("previewBox");
    if (!pre || pre.parentElement.querySelector(".preview-tools")) return;
    const tools = document.createElement("div");
    tools.className = "preview-tools";
    tools.innerHTML = `
      <button type="button" class="ghost" id="copyPreviewBtn" title="TXT másolása vágólapra">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
        Másolás
      </button>
      <button type="button" class="ghost" id="expandPreviewBtn" title="Nagyítás">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
        Teljes nézet
      </button>`;
    pre.parentElement.insertBefore(tools, pre);

    document.getElementById("copyPreviewBtn").addEventListener("click", async () => {
      const t = (pre.textContent || "").trim();
      if (!t || t.startsWith("A sikeres konvertálás")) {
        window.bankiToast?.("Még nincs TXT tartalom a vágólapra.", "warn"); return;
      }
      try { await navigator.clipboard.writeText(t); window.bankiToast?.("TXT a vágólapon.", "ok"); }
      catch {
        const ta = document.createElement("textarea");
        ta.value = t; document.body.appendChild(ta); ta.select();
        try { document.execCommand("copy"); window.bankiToast?.("TXT a vágólapon.", "ok"); }
        catch { window.bankiToast?.("Másolás nem sikerült.", "error"); }
        ta.remove();
      }
    });

    document.getElementById("expandPreviewBtn").addEventListener("click", () => {
      let dlg = document.getElementById("previewDialog");
      if (!dlg) {
        dlg = document.createElement("dialog");
        dlg.id = "previewDialog";
        dlg.className = "preview-dialog";
        dlg.innerHTML = `
          <div class="dialog-head">
            <h2>TXT teljes nézet</h2>
            <div style="display:flex;gap:8px;">
              <button type="button" class="secondary" id="copyPreviewBtn2">Másolás</button>
              <button type="button" class="secondary" id="closePreviewDlg">Bezárás</button>
            </div>
          </div>
          <div class="dialog-body"><pre id="previewBoxFull"></pre></div>`;
        document.body.appendChild(dlg);
        dlg.addEventListener("click", e => { if (e.target === dlg) dlg.close(); });
        dlg.addEventListener("cancel", () => dlg.close());
      }
      dlg.querySelector("#previewBoxFull").textContent = pre.textContent || "";
      dlg.querySelector("#closePreviewDlg").onclick = () => dlg.close();
      dlg.querySelector("#copyPreviewBtn2").onclick = () => document.getElementById("copyPreviewBtn").click();
      if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
    });
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
