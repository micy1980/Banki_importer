/* Banki UI - compact inline help icons for form fields. */
(function () {
  "use strict";

  const HELP = {
    ownAccountNumber: "Magyar 3x8 számlaszám kötőjelekkel, vagy HU IBAN. Magyar számlánál a banknév az első 8 számjegyből tölthető.",
    editAccountNumber: "Magyar 3x8 számlaszám vagy HU IBAN. Az IBAN MOD-97 algoritmussal ellenőrzött.",
    ownBankCountry: "A bank országkódja. Belföldi számlánál HU.",
    editBankCountry: "A bank országkódja. Belföldi számlánál HU.",
    ownCurrency: "ISO 4217 devizakód, például HUF, EUR vagy USD.",
    editCurrency: "ISO 4217 devizakód.",
    partnerCode: "Saját könyvelési vagy törzsadat-azonosító.",
    partnerName: "A partner neve.",
    partnerAccount: "Magyar 3x8 belföldi számlaszám.",
    partnerIban: "Nemzetközi IBAN számlaszám. Országonként eltérő hosszúságú lehet.",
    partnerSwift: "SWIFT/BIC: 8 vagy 11 karakter.",
    partnerCountry: "ISO országkód, például HU, DE vagy AT.",
    encoding: "A letöltött TXT karakterkódolása. Magyar banki importnál általában windows-1250.",
    identifierDate: "Ha az Excelben nincs azonosító, ebből a dátumból készül a generált azonosító.",
    bankSelect: "A választott bank szűri az elérhető importformátumokat.",
    formatSelect: "A banknál használható importformátum.",
  };

  function attach(id, text) {
    const input = document.getElementById(id);
    if (!input || input.dataset.helpAttached) return;
    const label = document.querySelector(`label[for="${id}"]`);
    if (!label) return;

    const button = document.createElement("button");
    button.type = "button";
    button.className = "field-help-btn";
    button.setAttribute("aria-label", `Súgó: ${label.textContent.trim()}`);
    button.title = text;
    button.textContent = "?";
    label.appendChild(button);
    input.dataset.helpAttached = "1";
  }

  function attachAll() {
    Object.entries(HELP).forEach(([id, text]) => attach(id, text));
  }

  if (document.readyState !== "loading") attachAll();
  else document.addEventListener("DOMContentLoaded", attachAll);
  document.addEventListener("click", event => {
    if (event.target?.id && /Btn$/.test(event.target.id)) setTimeout(attachAll, 60);
  });
})();
