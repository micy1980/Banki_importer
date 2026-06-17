/* Banki UI - Export summary screen.
   Adds an "Összegzés" button next to #convertBtn. Clicking it:
     1. Calls /api/convert (same payload as a download).
     2. Parses the returned TXT to count records and sum amounts.
     3. Shows a confirmation dialog with stats + a "Letöltés most" button.
   Pure additive: does not touch the existing convertBtn handler.
*/
(function () {
  "use strict";

  // PAYORD heuristics: amount lives in fixed columns on the "DET" / detail line.
  // We do not assume a specific column - instead we scan numeric tokens and
  // pick the largest plausible amount per record.
  function parseTxt(text) {
    const lines = text.split(/\r\n|\r|\n/).filter(l => l.length > 0);
    const records = lines.length;
    let totalAmount = 0;
    let partners = new Set();
    let hash = 0;
    for (const line of lines) {
      // FNV-1a 32-bit hash
      for (let i = 0; i < line.length; i++) {
        hash ^= line.charCodeAt(i);
        hash = (hash * 16777619) >>> 0;
      }
      // amount: longest run of digits >= 4 chars (in fillér / EDIFACT cents)
      const nums = line.match(/\d{4,}/g);
      if (nums && nums.length) {
        const n = parseInt(nums.sort((a, b) => b.length - a.length)[0], 10);
        if (!isNaN(n) && n < 1e15) totalAmount += n;
      }
      // partner: chars 200-300 region tends to hold name; use a slice
      const slice = line.slice(200, 260).trim();
      if (slice) partners.add(slice.slice(0, 40));
    }
    return {
      records,
      partners: partners.size,
      totalAmount: totalAmount / 100, // assume fillér -> Ft
      hash: hash.toString(16).padStart(8, "0").toUpperCase(),
      bytes: text.length,
    };
  }

  function fmtNum(n) {
    return new Intl.NumberFormat("hu-HU").format(Math.round(n));
  }

  function ensureDialog() {
    let dlg = document.getElementById("exportSummaryDialog");
    if (dlg) return dlg;
    dlg = document.createElement("dialog");
    dlg.id = "exportSummaryDialog";
    dlg.className = "preview-dialog";
    dlg.innerHTML = `
      <div class="dialog-head">
        <h2>Export összegzés</h2>
        <button type="button" class="secondary" data-close>Bezárás</button>
      </div>
      <div class="dialog-body">
        <p class="muted-small">Ellenőrizd az értékeket, majd indítsd a letöltést.</p>
        <div class="result-grid" id="exportSummaryGrid"></div>
        <div style="margin-top:14px;">
          <h3 style="font-size:13px;margin:0 0 8px;color:var(--ink-500);text-transform:uppercase;letter-spacing:.05em;">Első rekord</h3>
          <pre id="exportSummaryPreview" style="max-height:180px;overflow:auto;background:var(--surface-2);padding:10px;border-radius:8px;font-family:var(--font-mono);font-size:11px;"></pre>
        </div>
        <div class="account-actions" style="margin-top:16px;justify-content:flex-end;">
          <button type="button" class="secondary" data-close>Mégse</button>
          <button type="button" class="primary" id="exportSummaryDownload">Letöltés most</button>
        </div>
      </div>`;
    document.body.appendChild(dlg);
    dlg.addEventListener("click", e => {
      if (e.target === dlg || e.target.dataset.close !== undefined) dlg.close();
    });
    dlg.addEventListener("cancel", () => dlg.close());
    return dlg;
  }

  function renderGrid(stats) {
    const grid = document.getElementById("exportSummaryGrid");
    grid.innerHTML = `
      <div class="metric"><strong>${fmtNum(stats.records)}</strong><span>Rekord</span></div>
      <div class="metric"><strong>${fmtNum(stats.partners)}</strong><span>Egyedi partner</span></div>
      <div class="metric"><strong>${fmtNum(stats.totalAmount)} Ft</strong><span>Becsült összeg</span></div>
      <div class="metric"><strong>${stats.hash}</strong><span>Tartalom-hash</span></div>
      <div class="metric"><strong>${fmtNum(stats.bytes)} B</strong><span>Fájl méret</span></div>
    `;
  }

  function triggerDownload(text, filename, contentType) {
    const blob = new Blob([text], { type: contentType || "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function runSummary(btn) {
    const fileInput = document.getElementById("fileInput");
    if (!fileInput || !fileInput.files || !fileInput.files[0]) {
      window.bankiToast?.("Előbb tölts fel egy fájlt az Import dialógusban.", "warn");
      return;
    }
    btn.disabled = true;
    const originalLabel = btn.textContent;
    btn.textContent = "Összegzés készül…";
    try {
      // Reuse the same config that app.js would send. collectConfig() is local
      // to app.js; we reproduce by gathering the public element values.
      const config = (typeof window.collectConfig==="function") ? window.collectConfig() : null;
      const form = new FormData();
      form.append("file", fileInput.files[0]);
      if (config) form.append("config", JSON.stringify(config));
      else {
        // best-effort fallback: build minimal config from visible selects
        form.append("config", JSON.stringify({
          encoding: document.getElementById("encoding")?.value || "cp1250",
        }));
      }
      const res = await fetch("/api/convert", { method: "POST", body: form });
      if (!res.ok) {
        const ct = res.headers.get("content-type") || "";
        const data = ct.includes("application/json") ? await res.json() : { error: await res.text() };
        window.bankiToast?.(data.error || "Konvertálás sikertelen.", "error");
        return;
      }
      const buf = await res.arrayBuffer();
      const enc = document.getElementById("encoding")?.value || "cp1250";
      const label = enc === "cp1250" ? "windows-1250" : enc === "cp852" ? "ibm852" : "utf-8";
      let text = "";
      try { text = new TextDecoder(label).decode(buf); }
      catch { text = new TextDecoder("utf-8").decode(buf); }
      const stats = parseTxt(text);
      const dlg = ensureDialog();
      renderGrid(stats);
      document.getElementById("exportSummaryPreview").textContent =
        (text.split(/\r\n|\r|\n/).find(l => l.length) || "").slice(0, 800);

      // Save for diff feature
      try {
        const prev = localStorage.getItem("banki.lastExport.text");
        if (prev) localStorage.setItem("banki.prevExport.text", prev);
        localStorage.setItem("banki.lastExport.text", text);
        localStorage.setItem("banki.lastExport.stats", JSON.stringify(stats));
        localStorage.setItem("banki.lastExport.ts", new Date().toISOString());
      } catch (_) { /* quota */ }

      const filename = res.headers.get("x-filename") || "payord_import.txt";
      const ct = res.headers.get("content-type") || "text/plain";
      const dlBtn = document.getElementById("exportSummaryDownload");
      dlBtn.onclick = () => {
        triggerDownload(new Blob([buf], { type: ct }), filename, ct);
        // Use blob directly to preserve encoding
        const blob = new Blob([buf], { type: ct });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        dlg.close();
        window.bankiToast?.(`${stats.records} rekord letöltve.`, "ok");
      };
      if (typeof dlg.showModal === "function") dlg.showModal(); else dlg.setAttribute("open", "");
    } catch (e) {
      window.bankiToast?.("Összegzés hiba: " + (e.message || e), "error");
    } finally {
      btn.disabled = false;
      btn.textContent = originalLabel;
    }
  }

  function init() {
    const convertBtn = document.getElementById("convertBtn");
    if (!convertBtn || document.getElementById("summaryBtn")) return;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.id = "summaryBtn";
    btn.className = "secondary";
    btn.textContent = "Összegzés";
    btn.title = "Konvertálás előnézet letöltés nélkül - megnézheted, mi készül.";
    convertBtn.parentElement.insertBefore(btn, convertBtn);
    btn.addEventListener("click", () => runSummary(btn));
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
