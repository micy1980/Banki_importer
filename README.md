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
