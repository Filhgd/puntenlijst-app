# Puntenlijst-generator - handleiding voor gebruikers

**Universiteit Antwerpen - Faculteit Geneeskunde en Gezondheidswetenschappen**

Deze app leest de deliberatie-PDF's (Master verpleeg- en vroedkunde +
schakeljaar) en maakt er automatisch één overzichtelijk Excel-bestand van met
alle examenpunten, per afstudeerrichting, inclusief ingebouwde controle.

Ontwikkeld door Filip Haegdorens (filip.haegdorens@uantwerpen.be).

---

## 1. Downloaden

Je krijgt van Filip een van deze bestanden:

| Jouw computer | Bestand |
|---|---|
| Mac | `Puntenlijst-mac.zip` |
| Windows | `Puntenlijst.exe` |

Deze links wijzen altijd naar de nieuwste versie:

- Mac: <https://github.com/Filhgd/puntenlijst-app/releases/latest/download/Puntenlijst-mac.zip>
- Windows: <https://github.com/Filhgd/puntenlijst-app/releases/latest/download/Puntenlijst.exe>

Er is niets te installeren: het is één bestand dat je gewoon opent.
Bewaar het bv. op je Bureaublad of in je Documenten.

---

## 2. De app gebruiken

1. Open de app (zie hieronder voor de allereerste keer).
2. Sleep de **map** met de deliberatie-PDF's op het venster.
   Losse PDF-bestanden slepen kan ook, of kies ze via de knoppen
   "Kies een map…" / "Kies PDF-bestanden…".
3. Volg de voortgang in het logvenster. Na enkele seconden verschijnt
   "KLAAR" met een samenvatting.
4. Het Excel-bestand (`Puntenlijst_JJJJ-MM-DD_uumm.xlsx`) staat in dezelfde
   map als de PDF's. Klik op **"Open Excel"** om het meteen te openen, of
   **"Toon in map"** om het in Finder/Verkenner te zien.

Met de knop **"Info"** open je een venster met het versienummer, de
contactgegevens en de volledige versiegeschiedenis.

De app herkent zelf welk type PDF je aanlevert:

- **Deliberatielijsten** (diplomajaar) leiden tot de gewone tabbladen per richting.
  De app leest uit de kopregel of het om de eerste of de tweede zit gaat.
- **"Rapport Vaststelling Punten"** (niet-diplomajaar, met eerste en tweede zit)
  leidt tot de extra tabbladen "Master niet-diplomajaar" en
  "Schakeljaar niet-diplomajaar".

Je mag alle types gerust samen in één map zetten; ze komen automatisch op de
juiste tabbladen terecht en gebruiken dan dezelfde vaknamen.

### Eerste en tweede zit samen

Sleep je zowel de lijsten van de **eerste** als van de **tweede zit** in de app,
dan worden studenten samengevoegd op studentnummer: één rij per student, met
per vak drie kolommen (**1e**, **2e**, **eind**). De kolom "2e" wordt alleen
ingevuld bij vakken waar het punt effectief veranderd is, dus bij echte
herkansingen. De kolom **Zit** toont welke lijsten voor die student gevonden
zijn ("1+2", "1" of "2"). De controle gebeurt dan tegen de eindtoestand, met de
getallen uit de tweedezitlijst.

Lever je alléén de tweedezitlijst aan, dan blijft de gewone smalle opmaak
behouden: die lijst bevat immers al alle vakken met hun definitieve punt.

### Wat staat er in het Excel-bestand?

- Een tabblad per afstudeerrichting + een gecombineerd mastertabblad.
- Studenten in rijen, vakken in kolommen; punten op 20.
- Rood = tekort (punt < 10 of code zoals AFW/NG), groen = vrijstelling (G/VZP).
- Tabbladen **"Master niet-diplomajaar"** en **"Schakeljaar niet-diplomajaar"** (als je zulke
  PDF's aanlevert): per vak drie kolommen, namelijk **1e** (eerste zit),
  **2e** (tweede zit, geel) en **eind** (grijs; het resultaat van de tweede
  zit, of dat van de eerste als er geen tweede zit was). Rechts staan het
  gemiddelde van de eindpunten, het aantal tekorten, het aantal vakken en het
  aantal vakken met een tweede zit. Let op: deze rapporten bevatten geen
  controlegetallen, dus de automatische kruiscontrole geldt hier niet.
- Tabblad **Controle**: per student wordt nagekeken of het aantal gevonden
  vakken en tekorten klopt met wat in de PDF staat. Staat er ergens
  "CONTROLEER", kijk die student dan even na.
- Tabblad **Legende**: alle vakcodes met vaknaam, studiepunten en lector.

### Sorteren, filteren en kolommen verbergen

- Elk tabblad staat **gesorteerd van hoog naar laag op "Resultaat %"**. Waar de
  PDF geen percentage vermeldt (de tabbladen van het niet-diplomajaar) gebeurt
  dat op het gemiddelde eindpunt.
- De **filterknoppen staan al aan** op de kopregel. Klik op een pijltje om zelf
  te sorteren of te filteren, bijvoorbeeld op richting of op "CONTROLEER".
- Op de tabbladen met een eerste én tweede zit staan de kolommen **"1e" en
  "2e" bij het openen ingeklapt**: je ziet meteen het overzicht met enkel de
  eindresultaten. Klik op een **plusje** boven een kolom om dat ene vak open
  te klappen, of gebruik de knopjes **1** en **2** linksboven het werkblad:
  **2** klapt alles open, **1** klapt alles weer dicht.
- Bovenaan elke vakkolom staat de **vakcode én de vaknaam**. Beweeg met de muis
  over de kop voor de studiepunten en de lector.

---

## 3. Eerste keer openen op een Mac

De app is niet ondertekend met een (betalend) Apple-certificaat. macOS
blokkeert daarom standaard de eerste start. Dit is normaal en eenmalig.

**Stap voor stap:**

1. Dubbelklik op het zip-bestand; je krijgt `Puntenlijst.app`.
2. Dubbelklik op `Puntenlijst.app`. Je krijgt een melding zoals
   *"'Puntenlijst' kan niet worden geopend"* - klik op **Gereed** (nog niet
   op Verplaats naar prullenmand!).
3. Open **Systeeminstellingen** → **Privacy en beveiliging**.
4. Scrol helemaal naar beneden. Bij *Beveiliging* staat:
   *"'Puntenlijst' is geblokkeerd..."* - klik op **Open toch**
   (op oudere macOS: "Toch openen").
5. Bevestig met je wachtwoord of Touch ID en klik nogmaals **Open**.
6. Vanaf nu opent de app gewoon met dubbelklikken.

**Alternatief (oudere macOS-versies):** houd de Ctrl-toets ingedrukt, klik op
de app → **Open** → **Open**.

**Melding "is beschadigd en kan niet worden geopend"?** Dat is dezelfde
beveiliging in een ander jasje. Oplossing: open de app **Terminal**
(Launchpad → zoek "Terminal") en typ (pas het pad aan naar waar de app staat):

```
xattr -cr ~/Downloads/Puntenlijst.app
```

Druk op Enter en open de app daarna opnieuw.

---

## 4. Eerste keer openen op Windows

1. Dubbelklik op `Puntenlijst.exe`.
2. Windows SmartScreen kan een blauw venster tonen: *"Uw pc wordt beschermd"*.
3. Klik op **Meer informatie** en daarna op **Toch uitvoeren**.
4. Vanaf nu opent de app zonder vragen.

De eerste start kan 10-20 seconden duren (de app pakt zichzelf uit);
daarna gaat het sneller.

---

## 5. Veelgestelde vragen

**Er gebeurt niets als ik mijn map sleep.**
Controleer of er PDF-bestanden in de map zitten. De app verwerkt alleen
`.pdf`-bestanden.

**De app zegt "Geen studenten gevonden".**
Waarschijnlijk zijn dit geen deliberatie-PDF's, of zijn het gescande
afbeeldingen in plaats van originele PDF-exports.

**Het tabblad Controle toont "CONTROLEER" bij een student.**
Het aantal gevonden vakken of tekorten wijkt af van wat de PDF vermeldt.
Kijk die student handmatig na in de originele PDF. De kolom "Detail" zegt
precies wat er afwijkt.

**Waar komt het Excel-bestand terecht?**
In dezelfde map als de PDF's. Lukt schrijven daar niet (bv. netwerkmap),
dan op je Bureaublad.

**Mag ik het bestand hernoemen of verplaatsen?**
Ja, het is een gewoon Excel-bestand.

---

## 6. Problemen of vragen?

Mail Filip Haegdorens: **filip.haegdorens@uantwerpen.be**
Vermeld de versie (staat onderaan in het venster van de app) en, bij een
foutmelding, een schermafbeelding van het logvenster.

---

## 7. Versiegeschiedenis

Je vindt deze lijst ook in de app zelf, via de knop **"Info"**.

### Versie 1.3.2 (9 september 2026)

- Kolommen "1e" en "2e" zijn inklapbaar en staan bij het openen van het
  bestand meteen dicht: je ziet dus eerst het overzicht met enkel de
  eindresultaten. Met de plusjes open je één vak, met het knopje 2 linksboven
  open je alles.
- Elk tabblad is gesorteerd van hoog naar laag op "Resultaat %" (en op het
  gemiddelde eindpunt waar de PDF geen percentage geeft).
- De filterknoppen staan standaard aan op alle tabbladen.
- De kolomkop toont voortaan de vakcode én de vaknaam.
- Opgelost: in de rapporten van het niet-diplomajaar werden punten soms in de
  verkeerde kolom gezet wanneer het aantal studiepunten ontbrak. De punten
  worden nu toegewezen op basis van hun plaats in de PDF, zodat 1e en 2de zit
  altijd correct uit elkaar gehouden worden.
- Opgelost: op vervolgpagina's zonder kopregel viel de app terug op de oude,
  foutgevoelige herkenning. De kolomposities worden nu onthouden voor het hele
  rapport.
- Opgelost: een herkansing waarbij de student hetzelfde punt haalde, bleef
  onzichtbaar in de kolom "2e". Een tweede zit wordt nu herkend aan de
  examenperiode (bv. S01 in juni tegenover Z02 in september), niet enkel aan
  een gewijzigd punt.
- Tabbladen krijgen de drie kolommen enkel wanneer er voor die groep effectief
  een eerste én een tweede zit is; anders blijft de smalle opmaak behouden.

### Versie 1.2.0 (8 september 2026)

- Ondersteuning voor "Rapport Vaststelling Punten" (niet-diplomajaar): nieuwe
  tabbladen "Master niet-diplomajaar" en "Schakeljaar niet-diplomajaar", met
  per vak de kolommen 1e zit, 2de zit en eind.
- Vaknamen uit die rapporten sluiten aan bij de vaknamen van de
  deliberatielijsten, ook wanneer de PDF de vaknaam afkapt.
- Deliberatielijsten van de eerste en de tweede zit worden samengevoegd op
  studentnummer: één rij per student met 1e zit, 2de zit en eind. De kolom
  "2e" toont enkel de effectieve herkansingen.
- Nieuwe kolom "Zit" toont welke lijsten voor een student gevonden zijn
  ("1+2", "1" of "2"), plus een teller "# Herkansingen".
- De ingebouwde controle vergelijkt nu de eindtoestand met de getallen uit de
  meest recente lijst.
- Knop "Info" met de versiegeschiedenis toegevoegd.

### Versie 1.1.0

- Vuurwerk en een geluidje wanneer de verwerking gelukt is.
- Schuddend venster en een waarschuwingsgeluid bij een fout.

### Versie 1.0.0

- Eerste versie als app voor macOS en Windows, met drag & drop.
- Logo van de faculteit en een eigen app-icoon.
- Leest deliberatie-PDF's en maakt één Excel met alle examenpunten, een
  tabblad per afstudeerrichting, plus de tabbladen Controle en Legende.
