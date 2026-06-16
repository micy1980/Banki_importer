# Banki TXT konverter

Helyi webalkalmazas Excel `.xlsx` / `.xlsm` vagy CSV inputbol Erste EDIFACT fix szelessegu TXT importfajl keszitesehez.

## Tamogatott formatumok

- Erste Forint atutalas - EDIFACT PAYORD (DO), 941 karakter soronkent CR/LF-fel.
- Erste Deviza atutalas - EDIFACT PAYORD (IN), 1048 karakter soronkent CR/LF-fel.

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

1. Nyisd meg az `Import` menut.
2. Valaszd ki a bankot es a formatumot.
3. Tolts fel egy `.xlsx`, `.xlsm` vagy `.csv` fajlt.
4. Kattints a `Beolvasas` gombra.
5. Az `Oszlopok es alapertekek` reszben ellenorizd a hozzarendelest.
6. Zard be az import ablakot, majd kattints a `TXT letoltese` gombra.

Az app mindig az elso munkalapot olvassa, es kotelezoen az elso sort tekinti fejlecnek.

## Beallitasok

A bank/formatum, oszlop-hozzarendelesek es alapertekek tartosan mentodnek a `settings.json` fajlba. A fajl torolheto, az app inditaskor uj alapbeallitast hoz letre.

## Excel sablon

A `Excel sablon` gomb a kivalasztott formatumhoz keszit fejlecezett `.xlsx` mintat.

## Kimenet

- TXT fajl, alapertelmezett kodolas: `windows-1250`.
- A rekordokat `CR/LF` zarja.
- Az osszeg mezok balrol nullazva, tizedespont nelkul kerulnek a TXT-be.
