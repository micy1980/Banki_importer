# Web Components — custom element dialogok

## Állapot
- `dialog-base.js` — közös ős (`DialogBase`).
- `help-dialog.js` — működő minta (`<help-dialog>`).
- A többi 7 dialog HTML-je még az `app.py` HTML_PAGE-ben él. Fokozatosan
  átírható; a JS API (`openDialog("xxx")` / `closeDialog("xxx")`) nem változik.

## Migrációs lépés egy dialogra
1. Másold az `app.py`-ból a `<dialog id="xxx">…</dialog>` blokk markup-ját.
2. Hozz létre `static/js/components/xxx-dialog.js` fájlt:
   ```js
   class XxxDialog extends DialogBase {
     get template() { return `<dialog aria-labelledby="..."> ... </dialog>`; }
     onReady() { /* event hookok ide */ }
   }
   customElements.define("xxx-dialog", XxxDialog);
   ```
3. Az `app.py` HTML-ben cseréld a `<dialog id="xxx">…</dialog>` blokkot erre:
   ```html
   <xxx-dialog id="xxx"></xxx-dialog>
   ```
4. Vedd fel a script tag-et az `app.py`-ban (a `core-dom.js` UTÁN).

## Sorrend (javasolt, kockázat szerint nőve)
1. helpDialog ✅ (kész)
2. accountEditDialog (kis form)
3. partnerEditDialog (kis form)
4. accountsDialog (lista + form)
5. partnersDialog (lista + form)
6. companiesDialog (lista + form)
7. importDialog (legnagyobb, sok mapping)
