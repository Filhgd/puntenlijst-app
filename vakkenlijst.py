# -*- coding: utf-8 -*-
"""
vakkenlijst.py
==============

Officiële volgorde, modules, afkortingen en namen van de opleidingsonderdelen
van de Master verpleegkunde en vroedkunde en het schakelprogramma.

Bron: "Overzicht studiegidsnummers - afkortingen", academiejaar 2026-2027.

De kolommen in de Excel volgen deze volgorde. Vakken die hier niet in staan
(bv. van een andere opleiding, of uit een vorig programma) gaan niet verloren:
die komen achteraan onder de module "Overige opleidingsonderdelen".

Bijwerken voor een nieuw academiejaar: pas de regels hieronder aan. De volgorde
van de regels bepaalt de volgorde van de kolommen.
"""

# (vakcode, afkorting, module, jaar, studiepunten, vaknaam)
VAKKEN = [
    ('5001GENVEV', 'SMI1', 'Management en innovatie', 'Schakeljaar', '5',
     'Inleiding op management en innovatie in de gezondheidszorg'),
    ('5002GENVEV', 'SMI2', 'Management en innovatie', 'Schakeljaar', '5',
     'De ondernemer in de gezondheidszorg'),
    ('5003GENVEV', 'SMI3', 'Management en innovatie', 'Schakeljaar', '5',
     'Beleidsontwikkelingen in de gezondheidszorg in nationaal en internationaal perspectief'),
    ('5007GENVEV', 'SKV1', 'Klinische vorming', 'Schakeljaar', '5',
     'Verpleegkundige modellen en theorievorming, klinische aspecten van een zorgprogramma'),
    ('5008GENVEV', 'SKV2', 'Klinische vorming', 'Schakeljaar', '5',
     'Actuele ontwikkelingen en inzichten in de verloskunde'),
    ('5009GENVEV', 'SKV3', 'Klinische vorming', 'Schakeljaar', '5',
     'Somatische zorg'),
    ('5010GENVEV', 'SKV4', 'Klinische vorming', 'Schakeljaar', '5',
     'Geestelijke gezondheidszorg en ouderenzorg'),
    ('5011GENVEV', 'SAV1', 'Academische vorming', 'Schakeljaar', '5',
     'Persoonlijke ontwikkeling als geschoolde zorgverlener in de academische context'),
    ('5004GENVEV', 'SWO1', 'Wetenschappelijk onderzoek', 'Schakeljaar', '10',
     'Beginselen van onderzoeksmethodologie in verpleegkunde en vroedkunde'),
    ('5005GENVEV', 'SWO2', 'Wetenschappelijk onderzoek', 'Schakeljaar', '7',
     'Toepassingen van onderzoeksmethodologie in verpleegkunde en vroedkunde'),
    ('5006GENVEV', 'SWO3', 'Wetenschappelijk onderzoek', 'Schakeljaar', '3',
     'Masterproefvoorbereiding: het ontwerp van een onderzoeksprotocol'),
    ('2024GENVEV', 'MWO1', 'Gemeenschappelijke stam', 'Masterjaar', '4',
     'Systematisch literatuuronderzoek'),
    ('2025GENVEV', 'MWO2', 'Gemeenschappelijke stam', 'Masterjaar', '3',
     'Kwantitatieve data-analyse'),
    ('2026GENVEV', 'MWO3', 'Gemeenschappelijke stam', 'Masterjaar', '3',
     'Toepassen van kwalitatief onderzoek'),
    ('2020GENVEV', 'MMI1', 'Gemeenschappelijk (alle richtingen)', 'Masterjaar', '5',
     'Leiderschap als regie van zorg: concepten en vaardigheden'),
    ('2021GENVEV', 'MMI2', 'Gemeenschappelijk (alle richtingen)', 'Masterjaar', '5',
     'Leiderschap als regie van zorg: toepassing'),
    ('2022GENVEV', 'MKV1', 'Gemeenschappelijk (alle richtingen)', 'Masterjaar', '5',
     'De expert in het evidence based zorgproces'),
    ('2023GENVEV', 'MKV2', 'Gemeenschappelijk (alle richtingen)', 'Masterjaar', '5',
     'De professional als beheerder van kwaliteitszorg en patiëntveiligheid'),
    ('2030GENVEV', 'LGZ1', 'Leiderschap in gezondheid en zorg', 'Masterjaar', '5',
     'Leiderschap in gezondheid en zorg: ontwikkeling tot praktijkvoering'),
    ('2031GENVEV', 'LGZ2', 'Leiderschap in gezondheid en zorg', 'Masterjaar', '10',
     'Leiderschap in gezondheid en zorg: stage'),
    ('2037GENVEV', 'MP1 LGZ', 'Leiderschap in gezondheid en zorg', 'Masterjaar', '3',
     "Masterproef 'Leiderschap in gezondheid en zorg' - opzetten en opstarten"),
    ('2036GENVEV', 'MP2 LGZ', 'Leiderschap in gezondheid en zorg', 'Masterjaar', '12',
     "Masterproef 'Leiderschap in gezondheid en zorg'"),
    ('2050GENVEV', 'OGZ1', 'Onderzoeker in gezondheid en zorg', 'Masterjaar', '7',
     'Onderzoeker in gezondheid en zorg: stage'),
    ('2051GENVEV', 'OGZ2', 'Onderzoeker in gezondheid en zorg', 'Masterjaar', '4',
     'Actuele en complexe onderzoeksmethodologie'),
    ('2052GENVEV', 'OGZ3', 'Onderzoeker in gezondheid en zorg', 'Masterjaar', '4',
     'Verdieping in kwalitatief onderzoek'),
    ('2057GENVEV', 'MP1 OGZ', 'Onderzoeker in gezondheid en zorg', 'Masterjaar', '3',
     "Masterproef 'Onderzoeker in gezondheid en zorg' - opzetten en opstarten"),
    ('2056GENVEV', 'MP2 OGZ', 'Onderzoeker in gezondheid en zorg', 'Masterjaar', '12',
     "Masterproef 'Onderzoeker in gezondheid en zorg'"),
    ('2140GENVEV', 'VES1', 'Verpleegkundig specialist', 'Masterjaar', '9',
     'De verpleegkundig specialist als klinisch expert en behandelaar'),
    ('2041GENVEV', 'VES2', 'Verpleegkundig specialist', 'Masterjaar', '4',
     'Farmacologie'),
    ('2042GENVEV', 'VES3', 'Verpleegkundig specialist', 'Masterjaar', '4',
     'De verpleegkundig specialist als leider en mentor'),
    ('2043GENVEV', 'VES4', 'Verpleegkundig specialist', 'Masterjaar', '4',
     'De verpleegkundig specialist als autonome professional'),
    ('2045GENVEV', 'VES5', 'Verpleegkundig specialist', 'Masterjaar', '4',
     'De verpleegkundig specialist als coach en gezondheidspromotor'),
    ('2046GENVEV', 'VES6', 'Verpleegkundig specialist', 'Masterjaar', '10',
     'Verpleegkundig specialist: stage'),
    ('2047GENVEV', 'MP1 VES', 'Verpleegkundig specialist', 'Masterjaar', '3',
     "Masterproef 'Verpleegkundig specialist' - opzetten en opstarten"),
    ('2146GENVEV', 'MP2 VES', 'Verpleegkundig specialist', 'Masterjaar', '12',
     "Masterproef 'Verpleegkundig specialist'"),
    ('2060GENVEV', 'VRS1', 'Vroedvrouw specialist', 'Masterjaar', '5',
     'Klinische verdieping in de verloskunde'),
    ('2160GENVEV', 'VRS2', 'Vroedvrouw specialist', 'Masterjaar', '5',
     'De vroedvrouw specialist in een (inter)nationale en toekomstgerichte zorgcontext'),
    ('2161GENVEV', 'VRS3', 'Vroedvrouw specialist', 'Masterjaar', '10',
     'De vroedvrouw specialist: stage'),
    ('2067GENVEV', 'MP1 VRS', 'Vroedvrouw specialist', 'Masterjaar', '3',
     "Masterproef 'Vroedvrouw specialist' - opzetten en opstarten"),
    ('2066GENVEV', 'MP2 VRS', 'Vroedvrouw specialist', 'Masterjaar', '12',
     "Masterproef 'Vroedvrouw specialist'"),
]

MODULE_OVERIG = "Overige opleidingsonderdelen"

# Opzoektabellen, opgebouwd bij het importeren.
VOLGORDE = {r[0]: i for i, r in enumerate(VAKKEN)}
AFKORTING = {r[0]: r[1] for r in VAKKEN}
MODULE = {r[0]: r[2] for r in VAKKEN}
NAAM = {r[0]: r[5] for r in VAKKEN}


def sorteersleutel(code):
    """Positie van een vak in de officiële volgorde; onbekend gaat achteraan."""
    return (VOLGORDE.get(code, len(VAKKEN)), code or "")


def module_van(code):
    """Module waartoe het vak behoort."""
    return MODULE.get(code, MODULE_OVERIG)


def afkorting_van(code):
    """Afkorting van het vak, of een lege tekst als die niet bekend is."""
    return AFKORTING.get(code, "")


def zoek_op_naam(begin):
    """
    Zoek een vak waarvan de officiële naam met deze (mogelijk afgekapte)
    tekst begint. Geeft de vakcode terug, of None bij geen of meerdere
    treffers.
    """
    begin = (begin or "").strip().lower()
    if len(begin) < 8:
        return None
    treffers = [c for c, n in NAAM.items() if n.lower().startswith(begin)]
    return treffers[0] if len(treffers) == 1 else None
