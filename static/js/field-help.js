/* Banki UI - Inline help icons (?) for form fields. */
(function () {
  "use strict";
  const HELP = {
    ownAccountNumber: "Magyar 3×8 számlaszám (kötőjellel) vagy HU IBAN. A 8 jegyű prefixből töltjük a bank nevét.",
    editAccountNumber: "Magyar 3×8 számlaszám vagy HU IBAN. MOD-97 IBAN-ellenőrzés.",
    ownBankCountry: "Bank országkódja ISO 3166 alpha-2 szerint. Belföldi esetben HU.",
    editBankCountry: "Bank országkódja ISO 3166 alpha-2 szerint.",
    ownCurrency: "ISO 4217 devizakód (HUF, EUR, USD…).",
    editCurrency: "ISO 4217 devizakód.",
    partnerCode: "Saját könyvelési azonosító — tetszőleges, max. 35 karakter.",
    partnerName: "Cégszerű név, max. 70 karakter. EDIFACT-ban a NAD szegmensbe kerül.",
    partnerAccount: "Magyar 3×8 belföldi számlaszám. SWIFT mellett opcionális.",
    partnerIban: "IBAN formátum: országkód + 2 ellenőrző + max 30 karakter.",
    partnerSwift: "SWIFT/BIC: 8 vagy 11 karakter, A-Z és 0-9.",
    partnerCountry: "ISO 3166 alpha-2 (HU, DE, AT…).",
    encoding: "TXT karakterkészlet. Magyar banki importnál windows-1250 az alap. UTF-8 csak ha a bank elfogadja.",
    identifierDate: "Ha az Excelben nincs 14 jegyű azonosító, ebből generálunk: ÉÉÉÉHHNN + 6 jegyű sorszám.",
    bankSelect: "A kiválasztott bank szűri a formátum választékot.",
    formatSelect: "Banki fix-hosszúságú TXT séma. PAYORD DO = forint, IN = deviza.",
  };
  function attach(id, text) {
    const input = document.getElementById(id);
    if (!input || input.dataset.helpAttached) return;
    const label = document.querySelector(`label[for="${id}"]`);
    if (!label) return;
    const btn = document.createElement("button");
    btn.type = "button"; btn.className = "field-help-btn";
    btn.setAttribute("aria-label", `Súgó: ${label.textContent.trim()}`);
    btn.title = text;
    btn.textContent = "?";
    label.appendChild(document.createTextNode(" "));
    label.appendChild(btn);
    input.dataset.helpAttached = "1";
  }
  function attachAll() { Object.entries(HELP).forEach(([id, text]) => attach(id, text)); }
  if (document.readyState !== "loading") attachAll();
  else document.addEventListener("DOMContentLoaded", attachAll);
  // re-run on dialog open events
  document.addEventListener("click", e => {
    if (e.target?.id && /Btn$/.test(e.target.id)) setTimeout(attachAll, 60);
  });
})();
