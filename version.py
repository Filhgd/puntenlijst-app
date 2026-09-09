# -*- coding: utf-8 -*-
"""Versie-, ontwikkelaars- en wijzigingsinformatie van de Puntenlijst-generator."""

__version__ = "1.4.0"

APP_NAME = "Puntenlijst-generator"
DEVELOPER = "Filip Haegdorens"
DEVELOPER_EMAIL = "filip.haegdorens@uantwerpen.be"
ORGANISATION = "Universiteit Antwerpen - Faculteit Geneeskunde en Gezondheidswetenschappen"

# Versiegeschiedenis: (versie, datum, [wijzigingen]).
# Laat de datum leeg ("") als je die niet wil tonen.
# Nieuwste versie steeds bovenaan.
VERSION_HISTORY = [
    ("1.4.0", "9 september 2026", [
        "De vakken staan nu altijd in dezelfde, officiële volgorde, volgens "
        "het overzicht van de studiegidsnummers. Zo zien twee puntenlijsten "
        "er steeds hetzelfde uit.",
        "Nieuwe kopregel bovenaan met de module ('Wetenschappelijk onderzoek', "
        "'Klinische vorming', 'Gemeenschappelijke stam', de afstudeerrichting, "
        "...) die de bijhorende vakkolommen overspant.",
        "De kolomkop toont de afkorting, de vakcode en de vaknaam, "
        "bijvoorbeeld 'SWO1 - 5004GENVEV - Beginselen van onderzoeksmethodologie'.",
        "Vakken die niet in het overzicht staan (een andere opleiding, of een "
        "vak uit een vorig programma) gaan niet verloren: die komen achteraan "
        "onder de module 'Overige opleidingsonderdelen'.",
        "Vakken worden ook herkend aan hun officiële naam, zodat ze ook "
        "kloppen wanneer er voor die richting geen deliberatielijst is "
        "meegegeven.",
        "Het tabblad Legende toont nu ook de afkorting en de module, in "
        "dezelfde volgorde als de kolommen.",
        "Opgelost: afgekapte vaknamen werden bij elke student opnieuw "
        "aangevuld, waardoor het laatste stuk van de naam tientallen keren "
        "herhaald werd in de kolomkop.",
    ]),
    ("1.3.2", "9 september 2026", [
        "Kolommen '1e' en '2e' zijn inklapbaar en staan bij het openen van het "
        "bestand meteen dicht: je ziet dus eerst het overzicht met enkel de "
        "eindresultaten. Met de plusjes open je één vak, met het knopje 2 "
        "linksboven open je alles.",
        "Elk tabblad is gesorteerd van hoog naar laag op 'Resultaat %' "
        "(en op het gemiddelde eindpunt waar de PDF geen percentage geeft).",
        "De filterknoppen staan standaard aan op alle tabbladen, zodat je "
        "zelf kan sorteren en filteren.",
        "De kolomkop toont voortaan de vakcode én de vaknaam.",
        "Opgelost: in de rapporten van het niet-diplomajaar werden punten soms "
        "in de verkeerde kolom gezet wanneer het aantal studiepunten ontbrak. "
        "De punten worden nu toegewezen op basis van hun plaats in de PDF, "
        "zodat 1e en 2de zit altijd correct uit elkaar gehouden worden.",
        "Opgelost: op vervolgpagina's zonder kopregel viel de app terug op de "
        "oude, foutgevoelige herkenning. De kolomposities worden nu onthouden "
        "voor het hele rapport.",
        "Tabbladen krijgen de drie kolommen enkel wanneer er voor die groep "
        "effectief een eerste én een tweede zit is; anders blijft de smalle "
        "opmaak behouden.",
        "Opgelost: een herkansing waarbij de student hetzelfde punt haalde, "
        "bleef onzichtbaar in de kolom '2e'. Een tweede zit wordt nu herkend "
        "aan de examenperiode (bv. S01 in juni tegenover Z02 in september), "
        "niet enkel aan een gewijzigd punt.",
    ]),
    ("1.2.0", "8 september 2026", [
        "Ondersteuning voor 'Rapport Vaststelling Punten' (niet-diplomajaar): "
        "nieuwe tabbladen 'Master niet-diplomajaar' en 'Schakeljaar "
        "niet-diplomajaar' met per vak de kolommen 1e zit, 2de zit en eind.",
        "Vaknamen uit die rapporten sluiten aan bij de vaknamen van de "
        "deliberatielijsten, ook wanneer de PDF de vaknaam afkapt.",
        "Deliberatielijsten van de eerste en de tweede zit worden samengevoegd "
        "op studentnummer: één rij per student met 1e zit, 2de zit en eind. "
        "De kolom '2e' toont enkel de effectieve herkansingen.",
        "Nieuwe kolom 'Zit' toont welke lijsten voor een student gevonden zijn "
        "(1+2, 1 of 2), plus een teller '# Herkansingen'.",
        "De ingebouwde controle vergelijkt nu de eindtoestand met de getallen "
        "uit de meest recente lijst.",
        "Deze knop met de versiegeschiedenis toegevoegd.",
    ]),
    ("1.1.0", "", [
        "Vuurwerk en een geluidje wanneer de verwerking gelukt is.",
        "Schuddend venster en een waarschuwingsgeluid bij een fout.",
    ]),
    ("1.0.0", "", [
        "Eerste versie als app voor macOS en Windows, met drag & drop.",
        "Logo van de faculteit en een eigen app-icoon.",
        "Leest deliberatie-PDF's en maakt één Excel met alle examenpunten, "
        "een tabblad per afstudeerrichting, plus de tabbladen Controle en "
        "Legende.",
    ]),
]


def history_as_text():
    """Geef de versiegeschiedenis als leesbare tekst."""
    regels = []
    for versie, datum, items in VERSION_HISTORY:
        kop = f"Versie {versie}" + (f"  ({datum})" if datum else "")
        regels.append(kop)
        regels.append("-" * len(kop))
        for it in items:
            regels.append(f"  • {it}")
        regels.append("")
    return "\n".join(regels).rstrip()
