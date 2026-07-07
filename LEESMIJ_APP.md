# Puntenlijst-generator - app voor Mac en Windows

Een app met venster en drag & drop rond het bestaande `puntenlijst.py`-script.
Collega's zonder technische kennis slepen de map met deliberatie-PDF's op het
venster en krijgen het Excel-bestand terug. Geen installatie of Python nodig.

## Voor de gebruikers (collega's)

1. Download de app (link krijg je van Filip):
   - Mac: `Puntenlijst-mac.zip` → dubbelklik om uit te pakken → `Puntenlijst.app`
   - Windows: `Puntenlijst.exe`
2. Open de app.
3. Sleep de map met de deliberatie-PDF's op het venster
   (losse PDF's slepen kan ook, net als kiezen via de knoppen).
4. Klaar. Het Excel-bestand staat in dezelfde map als de PDF's.
   Klik op "Open Excel" om het meteen te bekijken.

### Eerste keer openen

- **Mac**: de app is niet ondertekend met een Apple-certificaat. Bij de eerste
  keer: rechtsklik op `Puntenlijst.app` → "Open" → "Open" bevestigen.
  Daarna opent hij gewoon met dubbelklikken. Krijg je "beschadigd en kan niet
  worden geopend"? Open dan Terminal en voer uit:
  `xattr -cr ~/Downloads/Puntenlijst.app` (pas het pad aan).
- **Windows**: SmartScreen kan waarschuwen bij de eerste keer.
  Klik "Meer informatie" → "Toch uitvoeren".

## Voor de beheerder (Filip): de apps bouwen

De apps worden automatisch gebouwd door GitHub (gratis), omdat een Mac-app
alleen op een Mac en een Windows-app alleen op Windows gebouwd kan worden.
Eenmalige setup, daarna is een nieuwe versie bouwen één klik.

### Eenmalig

1. Maak een account op https://github.com (als je er nog geen hebt).
2. Maak een nieuwe **private** repository, bv. `puntenlijst-app`.
3. Upload de inhoud van deze map `PuntenlijstApp` naar de repository.
   Dat kan via de webinterface: "Add file" → "Upload files". Let op dat de
   map `.github/workflows/` met `build.yml` mee wordt geüpload (via de
   webinterface kan dat door de hele map te slepen).

### Per nieuwe versie

1. Ga op GitHub naar het tabblad **Actions**.
2. Kies links "Build apps (macOS + Windows)" → knop **Run workflow**.
3. Na ± 5 minuten staan onder de afgeronde run twee downloads ("Artifacts"):
   `Puntenlijst-mac` en `Puntenlijst-windows`.
4. Stuur die bestanden naar je collega's (of publiceer een tag `v1.0` zodat
   ze onder "Releases" komen te staan met een vaste downloadlink).

## Lokaal draaien of testen (optioneel)

```
pip install -r requirements.txt
python puntenlijst_gui.py
```

De commandoregel-versie blijft ook werken:
`python puntenlijst_core.py "/pad/naar/map/met/pdfs"`

## Bestanden

- `puntenlijst_core.py` - alle parseer- en Excel-logica (identiek aan het
  oude script, maar als herbruikbare module met een `generate()`-functie).
- `puntenlijst_gui.py`  - het venster met drag & drop.
- `requirements.txt`    - benodigde bibliotheken.
- `.github/workflows/build.yml` - bouwt automatisch de Mac- en Windows-app.
