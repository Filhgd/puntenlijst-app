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
| Mac | `Puntenlijst-mac-v1.0.0.zip` |
| Windows | `Puntenlijst-v1.0.0.exe` |

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

### Wat staat er in het Excel-bestand?

- Een tabblad per afstudeerrichting + een gecombineerd mastertabblad.
- Studenten in rijen, vakken in kolommen; punten op 20.
- Rood = tekort (punt < 10 of code zoals AFW/NG), groen = vrijstelling (G/VZP).
- Tabblad **Controle**: per student wordt nagekeken of het aantal gevonden
  vakken en tekorten klopt met wat in de PDF staat. Staat er ergens
  "CONTROLEER", kijk die student dan even na.
- Tabblad **Legende**: alle vakcodes met vaknaam, studiepunten en lector.

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

1. Dubbelklik op `Puntenlijst-v1.0.0.exe`.
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
