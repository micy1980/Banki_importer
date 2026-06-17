# Lovable UI/UX audit prompt

Viselkedj ugy, mint egy senior product designer es senior frontend engineer, aki production-ready, akadalymentes, skalkozhato UI-rendszereket audital modern startup kornyezetben.

Feladatod: teljes UI/UX atvilagitas a Banki TXT konverter alkalmazason. Ez egy magyar nyelvu, banki importfajl-keszito webalkalmazas Excel/CSV beolvasassal, bankonkenti formatumvalasztassal, TXT exporttal, tobbceges mukodessel, sajat bankszamlakkal es partnerlistaval.

Fontos korlatozas:

- Ne modosits fajlt.
- Ne irj at kodot.
- Ne hozz letre commitot.
- Ne tervezz automatikus refaktort.
- Csak olvasd at es elemezd az alkalmazast, majd irj konkret, priorizalt javaslatokat.

Vizsgald meg kulonosen:

1. Informacios architektura es navigacio
   - Cegvalaszto, Import, Sajatat bankszamlak, Partnerlista, Beallitasok, Sugo.
   - Melyik funkcio legyen elsodleges, masodlagos vagy ritkan hasznalt?
   - Hol tul zsufolt a felulet, hol hianyzik kontextus?

2. Fo felhasznaloi folyamatok
   - Excel/CSV import.
   - Bank es formatum valasztas.
   - Beolvasas eredmenye, hibak es javitasi lehetosegek.
   - TXT letoltes.
   - Sajatat bankszamla rogzitese/importja.
   - Partner rogzitese/importja/szerkesztese.
   - Cegvaltas es ceghez kotott torzsadatok.

3. Komponensrendszer
   - Javasolj ujrahasznalhato komponenseket: Button, Select, Field, FilePicker, Dialog, Drawer, Toast, EmptyState, ErrorSummary, DataTable, StatusBadge, Toolbar, Tabs, FormSection.
   - Ird le a komponensek felelosseget es javasolt props/API tervet.
   - Jelezd, hol van duplikacio vagy inkonzisztens viselkedes.

4. Production-ready allapotok
   - Betoltesi allapotok.
   - Ures allapotok.
   - Hibak es validacios uzenetek.
   - Reszleges siker, import sorhibak, halozati hiba, MNB tabla frissitesi hiba.
   - Letoltes elotti tiltott vagy bizonytalan allapotok.

5. Akadalymentesseg
   - Billentyuzet-navigacio.
   - Dialog focus trap es visszafokuszalasa.
   - Label, aria-label, aria-describedby.
   - Kontraszt, focus ring, hiba osszekapcsolasa mezovel.
   - Kepernyoolvaso-baratsag.

6. Reszponziv viselkedes
   - Asztali szeles nezet.
   - Laptop meret.
   - Tablet.
   - Mobil.
   - Nagy tablazatok es horizontalis scroll.
   - Dialogok kisebb kepernyon.

7. Visual design
   - Tipografia, spacing, grid, hierarchia.
   - Banki/uzleti bizalmat erosito vizualis stilus.
   - Szinek es statuszjelolesek.
   - Gombok hierarchiaja.
   - Tablazatok olvashatosaga.

8. UX copy
   - Magyar nyelvu szovegek erthetosege.
   - Hibauzenetek: pontos ok + kovetkezo lepes.
   - Segitseg es beallitasok nyelvezete.

Kimenet formatuma:

1. Executive summary: 5-8 mondatban a legfontosabb megallapitasok.
2. Priorizalt teendolista:
   - P0: kritikus, hasznalatot akadalyozo gondok.
   - P1: fontos UX/UI javitasok.
   - P2: finomitasok es polish.
3. Kepernyonkenti audit:
   - Fo oldal
   - Import panel
   - Beolvasas eredmenye
   - TXT elonezet/export
   - Cegkezelo
   - Sajatat bankszamlak
   - Partnerlista
   - Beallitasok/Sugo
4. Komponensarchitektura-javaslat:
   - Komponensnev
   - Felelosseg
   - Javasolt props/API
   - Allapotok
   - Akadalymentessegi kovetelmenyek
5. Design system javaslat:
   - Tokenek: szinek, spacing, radius, shadow, font meretek
   - Gombhierarchia
   - Formmezok
   - Statuszok
   - Tablazatok
6. Implementacios backlog:
   - Feladat
   - Prioritas
   - Elfogadasi kriterium
   - Kockazat
7. Best practice lista:
   - Mit tartsunk meg?
   - Mit ne csinaljunk?
   - Milyen frontend szabalyokat vezessunk be?

Kerlek, a valasz legyen konkret, szakmai es megvalosithato. Ne altalanos design tanacsokat adj, hanem ehhez az alkalmazashoz illeszkedo, fejlesztoi backlogga alakithato javaslatokat.
