# -*- coding: utf-8 -*-
"""Versie-, ontwikkelaars- en wijzigingsinformatie van de Puntenlijst-generator."""

__version__ = "1.3.0"

APP_NAME = "Puntenlijst-generator"
DEVELOPER = "Filip Haegdorens"
DEVELOPER_EMAIL = "filip.haegdorens@uantwerpen.be"
ORGANISATION = "Universiteit Antwerpen - Faculteit Geneeskunde en Gezondheidswetenschappen"

# Versiegeschiedenis: (versie, datum, [wijzigingen]).
# Laat de datum leeg ("") als je die niet wil tonen.
# Nieuwste versie steeds bovenaan.
VERSION_HISTORY = [
    ("1.3.0", "8 september 2026", [
        "Kolommen '1e' en '2e' zijn nu inklapbaar: met de knopjes 1 en 2 "
        "linksboven in Excel zie je in één klik enkel de eindresultaten.",
        "Elk tabblad is gesorteerd van hoog naar laag op 'Resultaat %' "
        "(en op het gemiddelde eindpunt waar de PDF geen percentage geeft).",
        "De filterknoppen staan standaard aan op alle tabbladen, zodat je "
        "zelf kan sorteren en filteren.",
        "De kolomkop toont voortaan de vakcode én de vaknaam.",
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
