# Lovable UI/UX redesign implementation prompt

Viselkedj ugy, mint egy senior product designer es senior frontend engineer, aki eles uzleti alkalmazasok UI-rendszereit tervezi es implementalja.

## Cel

Modernizald a Banki TXT konverter teljes UI/UX-et ugy, hogy professzionalis, modern fintech / B2B SaaS munkafelület legyen belole. Ez nem marketing oldal, hanem napi hasznalatra szant produktivitas-eszkoz banki importfajlok keszitesehez.

## Jogosultsag

Ebben a feladatban modosithatsz fajlokat a repoban.

Engedelyezett:

- UI/UX, HTML, CSS es kliensoldali JavaScript modositasa.
- Komponensarchitektura kialakitasa.
- CSS kiszervezese kulon fajlba, ha ettol tisztabb lesz a kod.
- JavaScript kiszervezese kulon fajlba, ha ettol tisztabb lesz a kod.
- Dialogok, panelek, tablazatok, toolbarok, statuszok, ures allapotok es hibakezeles ujratervezese.
- Reszponziv viselkedes es akadalymentesseg javitasa.
- README frissitese, ha a hasznalat vagy a struktura valtozik.

Nem engedelyezett:

- Banki formatumlogika onkenyes megvaltoztatasa.
- TXT export rekordhosszok, mezopoziciok, validacios algoritmusok vagy banki szabalyok megvaltoztatasa, hacsak nem egyertelmuen UI-hibahoz kotodik.
- Runtime adatfajlok committolasa: `settings.json`, `companies.json`, `own_accounts.json`, `partners.json`, `bank_registry.json`.
- Titkok, tokenek, credentialok vagy szemelyes adatok felvetele a repoba.
- Olyan dependencia bevezetese, ami indokolatlanul neheziti a lokalis futtatast.

## Alkalmazas kontextus

Ez egy magyar nyelvu banki TXT importfajl-keszito webalkalmazas.

Fo funkciok:

- Excel / CSV feltoltes es beolvasas.
- Bank es formatum valasztas.
- HUF, Deviza es ahol dokumentalt, SEPA utalasok bankonkenti formatumban.
- TXT export fix hosszusagu banki importfajlokhoz.
- Tobbceges mukodes.
- Ceghez kotott sajat bankszamlak.
- Ceghez kotott partnerlista.
- Magyar bankszamlaszam 3x8 formatum es HU IBAN validacio.
- MNB hitelesito tabla automatikus frissitese.
- Magyar banknev felismerese prefix alapjan.

## Design irany

Talalj ki egy frissebb, letisztultabb, bizalmat epito fintech / B2B SaaS megjelenest.

Elvart karakter:

- modern, de nem harsany;
- banki es uzleti kornyezethez illo;
- adatgazdag, megis jol szkennelheto;
- gyors munkavegzesre optimalizalt;
- konzisztens statuszokkal es eros vizualis hierarchiaval;
- reszponziv, tablet/laptop meretben is jol hasznalhato.

Keruld:

- landing page jellegu hero design;
- tul dekorativ, marketinges megoldasok;
- zsufolt toolbar;
- egy szinten levo sok gomb;
- kartya-a-kartyaban szerkezet;
- bizonytalan disabled allapotok;
- csak szinnel jelolt hibak.

## Prioritasok

### P0

1. Modern fo layout
   - Legyen tiszta felso alkalmazas-header.
   - A cegvalaszto, aktiv bank/formatum es MNB tabla allapot legyen jol lathato.
   - A fo munkaterulet a beolvasas eredmenyere es TXT export folyamatra optimalizaljon.

2. Import workflow
   - Az import legyen egy jol strukturalt drawer vagy dialog.
   - Legyen egyertelmu sorrend: fajl -> bank/formatum -> beolvasas -> eredmeny -> TXT export.
   - A TXT letoltes maradjon kiemelt, elsodleges muvelet.

3. Hibakezeles
   - Legyen egységes ErrorSummary.
   - A hibauzenetek mondjak meg az okot es a kovetkezo lepest.
   - Import sorhibak legyenek jol olvashatoak.

4. Akadalymentes dialogok
   - Focus trap.
   - Esc zaras.
   - Fókusz visszaadas a megnyito gombra.
   - Label / aria-describedby / role status es role alert helyesen.

5. Reszponziv tablazatok
   - Nagy tablazatok desktopon sticky fejleccel.
   - Mobilon vagy kisebb kepernyon kartyas vagy jol scrollozhato nezet.

### P1

1. Design token rendszer
   - Szinek, spacing, radius, shadow, tipografia.
   - Legyen konzisztens gombhierarchia: primary, secondary, ghost, danger.

2. Ujrahasznalhato komponensmintak
   - Button
   - Select / Field
   - FilePicker
   - Dialog / Drawer
   - EmptyState
   - ErrorSummary
   - StatusBadge
   - DataTable
   - Toolbar
   - FormSection

3. Vizualis polish
   - Ikonok hasznalata ott, ahol segiti a gyors felismerest.
   - Jobb tipografia.
   - Jobb spacing.
   - Jobb statuszjelolesek.

### P2

1. TXT elonezet fejlesztese
   - Masolas vagolapra.
   - Kinyithato/nagyithato elonezet.

2. Nagy torzslistak kezelese
   - Kereses/szures a partner es bankszamla listaban.

3. Sugo es beallitasok finomitasa
   - Ne legyen zajos.
   - Kontextusban jelenjen meg, ahol segit.

## Elfogadasi kriteriumok

- Az app tovabbra is elindul `python app.py` paranccsal.
- A fo oldal betolt `http://127.0.0.1:8765` alatt.
- Az import panel/drawer megnyilik.
- A bank es formatum valasztas mukodik.
- Az Excel sablon letoltes megmarad.
- A TXT letoltes gomb allapotai erthetoek.
- A sajat bankszamlak es partnerek kulon ablakban/drawerben kezelhetoek.
- A dialogok billentyuzettel is hasznalhatoak.
- Mobil/laptop/desktop szelessegen nincs torott layout.
- Nincs runtime JSON adat committolva.
- Nincs token vagy credential a repoban.

## Vegso output

A modositas utan adj:

1. Rovid osszefoglalot, mi valtozott.
2. Kepernyonkenti felsorolast a fontos UI/UX javitasokrol.
3. Felsorolast, milyen fajlokat modositottal.
4. Validacios lepeseket, amiket futtattal.
5. Maradek javaslatokat kovetkezo iteraciora.

Kerlek, ne csak szepitsd a feluletet, hanem gondold ujra a munkafolyamat vizualis hierarchiajat is. A cel: kevesebb zaj, egyertelmubb kovetkezo lepes, modernebb fintech erzet, stabilabb production-ready felulet.
