# Banki TXT konverter

Helyi webalkalmazas banki importfajlok elkeszitesehez Excel vagy CSV inputbol. A felulet tobb ceggel, sajat bankszamlakkal, partnerlistaval es bankonkenti importformatumokkal dolgozik.

## Fobb funkciok

- Excel `.xlsx` / `.xlsm` es CSV beolvasas.
- Bank es formatum valasztas bankonkenti szuressel.
- TXT export fix hosszusagu banki importfajlokhoz.
- Cegvalaszto: az import kozos torzs, a sajat bankszamlak es partnerek ceghez kotodnek.
- Sajatat bankszamlak kezi rogzitese es importja.
- Partnerlista kezi rogzitessel, importtal es kulon szerkesztoablakkal.
- Magyar 3x8 bankszamlaszam es HU IBAN validacio.
- Magyar banknev automatikus felismerese az MNB hitelesito tablaja alapjan.
- Deviza rogzitese bankszamlaknal es partnereknel.
- Mozgathato dialogablakok.
- Excel sablon letoltes a valasztott importformatumhoz.

## Tamogatott bankok es utalastipusok

Az alkalmazas bankonkenti formatumlistat kezel. A cel a kovetkezo bankok sajat importformatumainak kezelese HUF, Deviza es ahol dokumentalt, SEPA utalashoz:

- Erste
- K&H
- MBH
- OTP
- Raiffeisen
- UniCredit

Megjegyzes: OTP-nel jelenleg a SEPA nincs kinalva, a deviza utalas az OTP dokumentacio szerinti `6. Deviza atutalasi megbizasok` formatumra epul. UniCredit SEPA csak akkor legyen aktivan hasznalhato, ha a konkret formatumleiras alapjan le van fejlesztve.

## Telepites

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Inditas

Windows alatt dupla kattintassal:

```text
start_payord_konverter.bat
```

Vagy terminalbol:

```powershell
python app.py
```

Utana bongeszoben:

```text
http://127.0.0.1:8765
```

## Hasznalat

1. Valaszd ki az aktiv ceget.
2. Nyisd meg az `Import` panelt.
3. Valaszd ki a bankot es a formatumot.
4. Tolts fel egy Excel vagy CSV fajlt.
5. Kattints a `Beolvasas` gombra.
6. Ellenorizd a beolvasas eredmenyet es a hibauzeneteket.
7. Kattints a `TXT letoltese` gombra.

Az app mindig az elso munkalapot olvassa, es kotelezoen az elso sort tekinti fejlecnek.

## Sajat bankszamlak

A `Sajat bankszamlak` menuben magyar bankszamlak rogzithetok kezzel vagy importtal. A magyar bankszamlaszam beiras kozben 3x8 formara alakul, es az alkalmazas ellenorzi a 9-7-3-1 sulyozasu ellenorzoszamokat.

HU IBAN eseten az IBAN MOD 97-10 ellenorzes is lefut, majd az app a belfoldi 24 szamjegyu szamlareszt is validalja.

## Hitelesito tabla

Az alkalmazas inditaskor megprobalja frissiteni az MNB hitelesito tablat:

```text
https://www.mnb.hu/letoltes/sht.xlsx
```

Ha a frissites sikerul, a felulet visszajelez. Ha nem sikerul, a korabban letoltott helyi tabla marad hasznalatban, es errol is statuszuzenet jelenik meg.

## Helyi adatfajlok

A futas kozben keletkezo helyi adatfajlok nincsenek Gitre szanva:

- `settings.json`
- `companies.json`
- `own_accounts.json`
- `partners.json`
- `bank_registry.json`

Ezek lokalis torzsadatot, beallitast vagy letoltott hitelesito tablat tartalmazhatnak.

## Kimenet

- TXT fajl, alapertelmezett kodolas: `windows-1250`.
- A rekordokat `CR/LF` zarja.
- Az osszeg mezok balrol nullazva, tizedespont nelkul kerulnek a TXT-be.


## UI / front-end architektura

A 2024-es UI/UX redesign ota a front-end forrasok kulon mappaban vannak, hogy az `app.py` ne tartalmazzon inline CSS-t es JS-t:

```
static/
  styles-base.css   # Az eredeti alaprajzu CSS (visszafele kompatibilis).
  tokens.css        # Modern design token reteg + sotet tema + iteracio 2 stilusok.
  app.js            # A teljes kliensoldali alkalmazaslogika.
  enhancements.js   # Toast rendszer, TXT masolas, teljes nezet, listaszuro, tema valto.
```

Az `app.py` a `/static/...` utvonalat ki is szolgalja, igy nincs szukseg kulon static serverre.

### Design token rendszer

A `static/tokens.css` `:root` szintu valtozokban definialja a szinpalettat, spacingot, radiust, shadow-t es tipografiat. Uj komponens irasakor mindig a tokeneket hasznald (`var(--brand-700)`, `var(--s-4)`, `var(--r-md)` stb.), ne hardkodolj szinerteket.

Komponens-szabalyok:

- Gombhierarchia: `primary` (oldalankent egy, banki piros gradiens), `secondary` (semleges), `ghost` (ikonok / mellekes muveletek), `danger` (torles).
- Statuszok: `status-pill.ok|warn|bad|info`, mindig szin + ikon-pont + szoveg egyutt.
- Dialogok: a `dialog` elem fokus-trap-pet es Esc-zarast az `app.js` mar kezeli, az `Import` dialogus jobb oldali drawer-kent jelenik meg.
- Tablazatok: `.sample-wrap` sticky fejleccel, scrollozhato, alternalo sorhatterekkel.

### Tema

A jobb felso sarokban levo hold ikon valt a sotet temara. A valasztas a `localStorage` `banki.theme` kulcsaba kerul. Sotet temahoz a `html.theme-dark` osztaly aktivalja az overrideokat a `tokens.css`-ben.

### Akadalymentesseg

- `*:focus-visible` lathato kek fokuszgyuru minden interaktiv elemen.
- `aria-live="polite"` a status uzeneten es a metric kartyakon, `role="alert"` a hibalistan.
- `prefers-reduced-motion` esetben a drawer animacio es minden tranzicio kikapcsol.
- Minden dialog billentyuzettel hasznalhato (Tab, Shift+Tab, Esc, Enter).

### Reszponziv viselkedes

- Desktop (>1080px): ket oszlopos workspace, balra `Beolvasas eredmenye`, jobbra `TXT elonezet`.
- Tablet (720-1080px): egy oszlopos workspace, sticky action rail, megmaradt drawer.
- Mobil (<720px): app header es action rail egymas ala kerul, metricek es import-grid egy oszlopban.

### Inline form validacio

Az `enhancements.js` az osszes editor dialog mezojere felrak egy `blur` figyelot:

- A `partnerName`, `editPartnerName`, `companyName`, `ownBankCountry`, `editBankCountry` kotelezo.
- A `partnerCountry` es `editPartnerCountry` ket nagybetus ISO orszagkodot var.
- A `partnerAccount`, `editPartnerAccount`, `ownAccountNumber`, `editAccountNumber` magyar 3x8 vagy IBAN formatumot var.
- A `partnerIban` es `editPartnerIban` IBAN szintaxist ellenoriz.
- A `partnerSwift` es `editPartnerSwift` 8 vagy 11 karakteres SWIFT/BIC-t var.

Ervenytelen ertek eseten a mezo `aria-invalid="true"`-ra valt es alatta megjelenik egy magyar nyelvu `field-error` magyarazat. Mentes elott a meglevo szerver oldali validacio tovabbra is lefut, ez az ellenorzes csak a felhasznaloi visszajelzest gyorsitja.

### Mobil kartyas tablanezet

A `partnersList`, `accountsList` es `companiesList` tablak 720 px alatt automatikusan kartyas listava alakulnak. Minden cellaba az `enhancements.js` `MutationObserver`-rel injektalja a fejlec szovegerol egy `data-label` attributumot, igy a kartyaban cimkezve jelennek meg az ertekek.

### Modularis kliensoldali architektura

Az `enhancements.js` monolitot kis, fokuszalt modulokra bontottuk a `static/js/` mappaba. Mindegyik egy IIFE, nincs build step, az `app.py` kulon `<script defer>` tagekkel toltibe oket:

| Fajl | Felelosseg |
| --- | --- |
| `toast.js` | `window.bankiToast(message, variant, ms)` - jobb also sarki ertesitesek. |
| `theme.js` | Sotet/vilagos tema valtas a fejlecbol, `localStorage`-ba mentve. |
| `preview.js` | TXT elonezet: vagolapra masolas + teljes nezet dialog. |
| `filter.js` | Partner / bankszamla / ceg dialog kereso mezo (kliens szuro). |
| `mobile-tables.js` | Tablak kartyas megjelenitese 720 px alatt (`data-label` injekcio). |
| `validation.js` | Mezoszintu inline validacio kotelezoseg, IBAN, BIC, magyar szamla, ISO ket-betus orszag szabalyokra. |
| `combobox.js` | A cegvalaszto `<select>`-hez hozzaillesztett kereso (csak >=8 cegnel jelenik meg). |
| `registry-retry.js` | Az MNB tabla pill kattinthato, `/api/bank-registry`-t ujrahivja, toast visszajelzes. |

Ha uj UX modult adsz hozza:

1. Tedd a forrast `static/js/<feature>.js`-be, IIFE-ben.
2. Vedd fel az `app.py` `HTML_PAGE` template `<script defer src="/static/js/<feature>.js"></script>` listajaba a tobbi modul mellol.
3. Globalisokat csak `window.banki*` nevterben tegy elerhetove (`bankiToast` mintajara).
4. Soha ne valtoztass a meglevo elem ID-kon, amiket az `app.js` hasznal - az ID szerzodes.

### MNB tabla ujraprobalkozas

A fejlec MNB pill `role="button"`-na valt. Kattintasra vagy Enter/Space lenyomasra fetcheli a `/api/bank-registry` vegpontot, frissiti a pillot es toastot dob. Az MNB letoltes elinditasahoz az app indulasakor lefuto httplogika a backenden valtozatlan; ez a retry csak a cache-elt allapot ujraolvasasat es ujrarendelesedt biztositja az aktualis valaszra. Tenyleges MNB ujraletoltes az alkalmazas ujraintidasakor tortenik.

### Cegvalaszto kereso

8 cegnel tobbnel a fejleckombobox melle automatikusan bekerul egy szabad szoveg kereso, amely begepelve szuri a `<select>` opcioit (hidden + disabled), es az elso talalatot ki is jeloli. Enter / Le nyil atadja a fokuszt a `<select>`-re. Kevesebb cegnel a kereso elrejtve marad.

## Smoke teszt

A `tests/test_smoke.py` egy stdlib alapu smoke teszt (`unittest` + `urllib`):

- elinditja az `app.py`-t,
- megvarja a 8765-os port figyelesi allapotat,
- ellenorzi az index HTML-t a redesign markereire,
- staticus fajlok kiszolgalasara HTTP 200-at vart,
- path traversal probat (`/static/../app.py`) 404 / 400-ra,
- `/api/settings`, `/api/companies`, `/api/bank-registry` JSON valaszat,
- `/template.xlsx` XLSX (PK header) letoltest.

Futtatas:

```powershell
python -m unittest tests.test_smoke -v
```

Vagy egyenesen:

```powershell
python tests\test_smoke.py
```

A teszt szubprocesszel inditja az appot, igy nem szabad mar futnia masik peldannyak a 8765-os porton.

## Iteration 5 - finishing touches

| # | Feature | Where |
|---|---------|-------|
| 1 | Real MNB re-fetch on pill click (POST `/api/bank-registry/refresh`) | `static/js/registry-retry.js` |
| 2 | Optional esbuild bundle (no runtime dep) | `tools/build.mjs` |
| 3 | Export summary screen with record count, partner count, total, hash | `static/js/export-summary.js` |
| 4 | Playwright E2E suite | `tests/playwright/` |
| 5 | `partnersDialog` / `accountsDialog` layout polish (sticky head, card density, responsive grid) | `static/tokens.css` |
| 6 | MNB pill tooltip showing `updated_at` + `row_count` + source URL | `static/js/registry-retry.js` |
| 7 | TXT diff viewer comparing previous and current export | `static/js/diff-view.js` |

### How the summary + diff hand off
1. User clicks **Összegzés** instead of **TXT letöltése**.
2. The browser calls `/api/convert` (same payload, same backend code path), parses the returned TXT in-memory to compute stats, and shows a confirmation dialog with **Letöltés most**.
3. The TXT is cached in `localStorage` as the "last export". The previous "last export" rotates into `prevExport`.
4. After two successful exports, a **Diff** button appears in the TXT preview toolbar, opening the line-level diff dialog.

### Optional bundle
The eight `static/js/*.js` modules ship as individual files (no build step required). For production builds:
```bash
npm i -D esbuild
node tools/build.mjs
```
then swap the eight `<script>` tags in `app.py` for the single bundle (see the script header for the exact line).

### Running E2E
See `tests/playwright/README.md`. The legacy stdlib smoke tests in `tests/test_smoke.py` continue to work and require no extra tooling.

## Iteration 6 - 14 új feature egyszerre

### Backend
| # | Feature | Endpoint / fájl |
|---|---------|------------------|
| 1 | Undo (törlés → kuka → helyreállítás) | `/api/{partners,accounts,companies}/{delete,restore}` + `_trash.json` |
| 3 | Audit log JSONL | `/api/audit` + `_audit.jsonl` |
| 11 | PWA — manifest, service worker, ikonok | `/manifest.webmanifest`, `/sw.js`, `/static/icon-{192,512}.png` |
| 14 | MNB tábla óránkénti háttér-frissítés | `background_registry_refresh()` thread, 24 órás ciklus |
| 15 | Docker + healthcheck | `Dockerfile`, `docker-compose.yml`, `/healthz` |

### Frontend (additív JS modulok `static/js/`)
| # | Feature | Modul |
|---|---------|-------|
| 1 | Undo toast „Visszavonás" gombbal (fetch-wrapper) | `undo.js` |
| 2 | Bulk műveletek (checkbox + tömeges törlés + CSV export) | `bulk-ops.js` |
| 3 | Audit napló dialog | `audit-log.js` |
| 5 | Mentett import-profilok (`localStorage`) | `import-profiles.js` |
| 6 | Onboarding wizard (első indításnál, ha nincs cég) | `onboarding.js` |
| 7 | Billentyűparancsok (`Ctrl+I` / `Ctrl+Enter` / `Ctrl+K`) + command palette | `shortcuts.js` |
| 8 | Mező-szintű súgó ikonok (`?`) | `field-help.js` |
| 9 | Drag & drop dropzone | `dropzone.js` |
| 10 | Recent files chipek | `recent-files.js` |
| 11 | PWA service worker regisztráció (csak production) | `pwa-register.js` |
| 12 | i18n váltó (HU/EN) | `i18n.js` |
| 13 | Hibasor → minta táblázat ugrás | `error-row-jump.js` |

### Roadmap (még nincs)
- **#4 Több cég párhuzamos import egy Excelből** — `convert_records()` mélyebb refaktorát igényli (cég-kulcs detektálás, per-cég kimenet, ZIP letöltés). Külön iterációban.

### Docker használat
```bash
docker compose up -d
# app: http://localhost:8765
# health: http://localhost:8765/healthz
```

### Billentyűparancsok
| Kombó | Akció |
|-------|-------|
| `Ctrl/Cmd + I` | Import dialog |
| `Ctrl/Cmd + Enter` | TXT letöltése |
| `Ctrl/Cmd + K` | Command palette |
| `Esc` | Bezárás (natív dialog) |

### PWA telepítés
Production üzemelésnél (HTTPS, nem localhost) a böngésző felajánlja a „Telepítés" opciót. Az app offline is működik a már látogatott oldalakra; `/api/*` hívások mindig hálózatot igényelnek.

## Iteration 7 — robusztusság & DevEx

- **Naplózás:** `print()` → Python `logging`. Szint: `BANKI_LOG_LEVEL=DEBUG|INFO|WARNING` (alap: `INFO`).
- **MNB frissítés intervalluma:** `BANKI_MNB_REFRESH_HOURS=24` (alap: 24 óra).
- **Audit log retenció:** `BANKI_AUDIT_MAX_LINES=20000` (alap). Túl hosszú napló esetén a régebbi sorok automatikusan `_audit.jsonl.1` fájlba kerülnek.
- **JSON séma verziózás:** `_schema_version.json` követi a perzisztens fájlok sémáját. Induláskor `run_schema_migrations()` lefuttatja a szükséges, idempotens migrációkat (jelenlegi verzió: **v2**). Új migrációt a `run_schema_migrations` `if v < N:` blokkjával lehet hozzáadni, majd `CURRENT_SCHEMA_VERSION` növelésével.
- **GitHub Actions CI:** `.github/workflows/ci.yml` Python 3.11/3.12/3.13 mátrixon syntax + pytest, illetve `bundle` job ami legyártja a `enhancements.bundle.js`-t.
- **JS bundler véglegesítve:** `tools/build.mjs` mostantól mind a 22 modult ismeri (a `static/js/` aktuális tartalma alapján). Használat:
  ```
  npm i -D esbuild
  node tools/build.mjs
  ```
  Majd `app.py`-ban a több `<script defer src="/static/js/*.js">` sor cserélhető egyre:
  `<script defer src="/static/dist/enhancements.bundle.js"></script>`.

> Auth, HTTPS, rate-limit szándékosan nincs — LAN + egyfelhasználós üzemmód.

## Iteráció 8 — backup, pre-commit, kód-struktúra

### Új végpont
- `GET /api/backup` — ZIP-be csomagol minden adatfájlt (`settings.json`, `own_accounts.json`,
  `companies.json`, `partners.json`, `_trash.json`, `_audit.jsonl`, `_schema_version.json`,
  `bank_registry.json`) + `_backup_manifest.json` (létrehozás ideje, séma verzió, fájlok).
  Fejlécben **Mentés (ZIP)** gomb is elérhető.

### Pre-commit hookok
```bash
pip install pre-commit
pre-commit install
pre-commit run --all-files   # első tisztítás
```
Tartalom: `ruff` (lint + format), `eslint` (flat config), `prettier`,
+ alap higiénia (trailing whitespace, EOF, JSON/YAML check, nagy fájl figyelmeztetés).
Konfig: `.pre-commit-config.yaml`, `eslint.config.mjs`, `.prettierrc.json`, `pyproject.toml`.

### Kód-struktúra változás
- `static/app.js`: **1242 → 1114 sor** (~10% csökkenés). A közös DOM helperek
  (`el`, `openDialog`, `closeDialog`, `focusableElements`, `setupDialogA11y`,
  `escapeHtml`, `normalise`, `optionList`, `emptyState`, `loadingRows`,
  `setButtonLoading`, `renderListState`, `fetchJson`) átkerültek
  `static/js/core-dom.js`-be. Az app.py ezt **app.js ELŐTT** tölti be `<script defer>`-rel,
  így a globális nevek változatlanok — nulla hívás-oldali változtatás.
- A bundler (`tools/build.mjs`) ORDER tömbje is frissült: `core-dom.js` az első.

### Web Components scaffold
- `static/js/components/` mappa, `help-dialog.js` minta + `README.md` konvenció.
- Cél: a 8 `<dialog>` fokozatos átírása custom element-ekké, fájlonként
  (nem big-bang). A `openDialog("xxx")` API kompatibilis marad, mert a custom
  element belül `<dialog>`-ot hordoz.

### Elhalasztva
- **TypeScript átállás**: opcionális volt, jelenleg vanilla JS marad. Ha jön rá igény,
  `static/js/*.js` → `*.ts`, `tsc --noEmit` a CI-ban, és az esbuild loader `ts`-re kapcsol.

## Iteráció 9 — modul-split + Web Component scaffold

### app.js szétszedése
**1114 → 478 sor** (~57% csökkenés). Új modulok a `static/js/` alatt:
- `accounts.js` (222 sor) — saját bankszámlák CRUD + import + validáció
- `partners.js` (171 sor) — partner CRUD + edit + BIC autofill
- `registry.js` (47 sor) — MNB bank-regisztrációs tábla + BIC lookup
- `convert.js` (166 sor) — config gyűjtés + PAYORD TXT generálás. Az event listener
  blokk `DOMContentLoaded`-be csomagolva, hogy az `app.js`-ben lévő globálisok
  (pl. `populateFormats`) már létezzenek mire fut.
- `dialogs.js` (35 sor) — drag-and-drop modális ablakok

Töltési sorrend (defer): `core-dom.js → components/* → accounts → partners → registry → convert → dialogs → app.js → enhancement modulok`.

**Megj.:** A megosztott állapot (`accountsState`, `currentInspect`, stb.) `let` helyett
`var`-ra változott `app.js`-ben, hogy a többi `<script defer>` is lássa
(classic script: top-level `var` globális, `let`/`const` script-scoped).
Hasonló okból a `core-dom.js` `const`-jai explicit `window.x = x` formában is publikálva vannak.

### Web Components alap
- `static/js/components/dialog-base.js` — közös ős (`class extends HTMLElement`).
- `static/js/components/help-dialog.js` — kész minta (`<help-dialog>`).
- `static/js/components/README.md` — konvenció + migrációs lépés-lista (7 dialógra).

A statikus fájlkiszolgáló most már 3 szintű útvonalat enged (`/static/js/components/x.js`).

### TypeScript átállás (még nem)
Vanilla JS marad. Ha jön rá igény: `static/js/*.js` → `*.ts`, esbuild loader `ts`-re,
`tsc --noEmit` a CI-ban.
