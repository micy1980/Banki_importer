/* Banki UI - Inline field-level validation for editor dialogs.
   Server-side validation remains the source of truth at save time. */
(function () {
  "use strict";
  const HU_ACC = /^\d{8}-\d{8}(-\d{8})?$|^\d{16}$|^\d{24}$/;
  const IBAN   = /^[A-Z]{2}\d{2}[A-Z0-9]{10,30}$/;
  const SWIFT  = /^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$/;
  const ISO2   = /^[A-Z]{2}$/;
  const RULES = {
    ownBankCountry:   { required:true,  hint:"Ország kötelező." },
    ownAccountNumber: { required:true,  test:v=>HU_ACC.test(v)||IBAN.test(v.replace(/\s+/g,"")),
                        hint:"Magyar 3x8 vagy érvényes IBAN szükséges." },
    editBankCountry:   { required:true, hint:"Ország kötelező." },
    editAccountNumber: { required:true, test:v=>HU_ACC.test(v)||IBAN.test(v.replace(/\s+/g,"")),
                         hint:"Magyar 3x8 vagy érvényes IBAN szükséges." },
    partnerName:    { required:true, hint:"A partner neve kötelező." },
    partnerCountry: { required:true, test:v=>ISO2.test(v.toUpperCase()),
                      hint:"Két nagybetűs ISO ország kód (pl. HU, DE, AT)." },
    partnerAccount: { test:v=>!v||HU_ACC.test(v),
                      hint:"Magyar 3x8 formátum: 12345678-12345678[-12345678]" },
    partnerIban:    { test:v=>!v||IBAN.test(v.replace(/\s+/g,"").toUpperCase()),
                      hint:"Érvényes IBAN: országkód + 2 jegy + max 30 alfanumerikus." },
    partnerSwift:   { test:v=>!v||SWIFT.test(v.toUpperCase()),
                      hint:"SWIFT/BIC: 8 vagy 11 karakter, csak nagybetű és szám." },
    editPartnerName:    { required:true, hint:"A partner neve kötelező." },
    editPartnerCountry: { required:true, test:v=>ISO2.test(v.toUpperCase()), hint:"Két nagybetűs ISO ország kód." },
    editPartnerAccount: { test:v=>!v||HU_ACC.test(v), hint:"Magyar 3x8 formátum." },
    editPartnerIban:    { test:v=>!v||IBAN.test(v.replace(/\s+/g,"").toUpperCase()), hint:"Érvényes IBAN szükséges." },
    editPartnerSwift:   { test:v=>!v||SWIFT.test(v.toUpperCase()), hint:"SWIFT/BIC formátum: 8 vagy 11 karakter." },
    companyName: { required:true, hint:"A cég neve kötelező." },
  };
  function wire(id, rule) {
    const f = document.getElementById(id); if (!f || f.dataset.validatorWired === "1") return;
    f.dataset.validatorWired = "1";
    const show = msg => {
      f.setAttribute("aria-invalid","true");
      let err = f.parentElement.querySelector(`[data-err-for="${id}"]`);
      if (!err) {
        err = document.createElement("div");
        err.className = "field-error"; err.setAttribute("data-err-for", id); err.id = id + "Err";
        f.setAttribute("aria-describedby", err.id);
        f.parentElement.appendChild(err);
      }
      err.textContent = msg;
    };
    const clear = () => {
      f.removeAttribute("aria-invalid");
      f.parentElement.querySelector(`[data-err-for="${id}"]`)?.remove();
    };
    const validate = () => {
      const v = (f.value || "").trim();
      if (rule.required && !v) { show(rule.hint || "Kötelező mező."); return false; }
      if (v && rule.test && !rule.test(v)) { show(rule.hint || "Érvénytelen érték."); return false; }
      clear(); return true;
    };
    f.addEventListener("blur", validate);
    f.addEventListener("input", () => { if (f.getAttribute("aria-invalid") === "true") validate(); });
  }
  function init() { Object.entries(RULES).forEach(([id, r]) => wire(id, r)); }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
