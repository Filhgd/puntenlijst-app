#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
puntenlijst_core.py
===================

Kernlogica van de Puntenlijst-generator: leest deliberatie-PDF's en bouwt
één Excel-bestand met alle examenpunten (op 20), gegroepeerd per
opleiding/afstudeerrichting.

Deze module bevat GEEN gebruikersinterface. Ze wordt gebruikt door:
  * puntenlijst_gui.py  - de grafische app (drag & drop)
  * de commandoregel:    python3 puntenlijst_core.py [map]
"""

import os
import re
import sys
import glob
from datetime import datetime

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# 1. Reguliere expressies om de tekst te ontleden
# ---------------------------------------------------------------------------

# Studenthoofding, bv: "20230477 Maes Lisa MASA - 0 Berekening: ..."
STUDENT_RE = re.compile(
    r"^(?P<nr>\d{8})\s+(?P<name>.+?)\s+(?P<status>[A-Z]{3,5})\s+-\s+\d"
)

# Planregel, bv: "Programma: M0029 - ... Plan: M0029004 - Ma VPVR: vroedvrouw spec"
PLAN_RE = re.compile(r"Plan:\s*(?P<plancode>\S+)\s*-\s*(?P<planname>.+?)\s*$")

# Vakregel-herkenning: begint met een vakcode en eindigt op
#   <PUNT> <OPLEIDINGSCODE> <aanbodsessie 4 cijfers> <volgnummer>
CODE_RE = re.compile(r"^\d{3,}[A-Z]{2,}\b")
TAIL_RE = re.compile(
    r"\s(?P<grade>\S+)\s+"
    r"(?P<tag>[A-Za-z0-9][A-Za-z0-9.\-]*)\s+"
    r"(?P<offer>\d{4})\s+\d+\s*$"
)
HEAD_RE = re.compile(
    r"^(?P<code>\S+)\s+(?P<name>.+?)\s+"
    r"(?P<sp>\d{1,2},\d{2})(?P<lect>.+?)\s+(?P<period>[SZ]\d{2})\s*$"
)

CREDIT_CODES = {"G", "VZP"}

EXAMENS_RE = re.compile(r"Totaal aantal examens:\s*(\d+)")
TEKORTEN_RE = re.compile(r"Aantal tekorten:\s*(\d+)")
RESULTAAT_RE = re.compile(r"Behaald resultaat:\s*(\d+)\s*%")
LIJST_RE = re.compile(r"\b([123])\s*-\s*(Witte|Grijze|Zwarte)\s+lijst", re.IGNORECASE)
BEOORDELING_RE = re.compile(r"\b([A-Z]{2,4})\s*-\s*([A-Z][A-Z ]{3,})\s*$")

PASS_THRESHOLD = 10  # punten < 10/20 worden gemarkeerd als tekort


# ---------------------------------------------------------------------------
# 2. Hulp: bepaal richting / jaar uit de planregel
# ---------------------------------------------------------------------------
TRACK_NAMES = {
    "M0029001": "Master - Onderzoeker in gezondheid en zorg",
    "M0029002": "Master - Leiderschap in gezondheid en zorg",
    "M0029003": "Master - Verpleegkundig specialist",
    "M0029004": "Master - Vroedvrouw specialist",
    "S0017000": "Schakeljaar - Verpleeg- en vroedkunde",
}

TRACK_SHORT = {
    "M0029001": "OGZ - Onderzoeker",
    "M0029002": "LGZ - Leiderschap",
    "M0029003": "VES - Verpleegk. spec.",
    "M0029004": "VRS - Vroedvrouw spec.",
    "S0017000": "Schakeljaar",
}


def classify_track(plancode, planname, status):
    """Geef (jaar, richtingsnaam, korte_naam) terug op basis van de planinfo."""
    plancode = (plancode or "").strip()
    if plancode.startswith("S0017") or status == "SPVP":
        jaar = "Schakeljaar"
    elif plancode.startswith("M0029") or status == "MASA":
        jaar = "Masterjaar"
    else:
        jaar = "Onbekend"
    naam = TRACK_NAMES.get(plancode)
    if not naam:
        naam = f"{jaar} - {planname}".strip(" -") if planname else (plancode or "Onbekende opleiding")
    kort = TRACK_SHORT.get(plancode) or (planname or naam)
    return jaar, naam, kort


def grade_to_value(raw):
    """Zet een ruwe puntentoken om. Geeft (weergave, numerieke_waarde_of_None)."""
    raw = raw.strip()
    if raw.isdigit():
        return str(int(raw)), int(raw)
    return raw, None


def is_deficit(disp, val):
    if val is not None:
        return val < PASS_THRESHOLD
    return disp not in CREDIT_CODES


def is_credit(disp, val):
    return val is None and disp in CREDIT_CODES


# ---------------------------------------------------------------------------
# 3. Eén PDF ontleden -> lijst van studentdicts
# ---------------------------------------------------------------------------
def parse_pdf(path, courses_registry, problems):
    students = []
    cur = None
    last_code = None
    in_courses = False

    def close():
        nonlocal cur
        if cur is not None:
            students.append(cur)
        cur = None

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                line = line.rstrip()
                if not line.strip():
                    continue

                m = STUDENT_RE.match(line)
                if m:
                    close()
                    cur = {
                        "nr": m.group("nr"),
                        "naam": m.group("name").strip(),
                        "status": m.group("status"),
                        "plancode": "",
                        "planname": "",
                        "grades": {},
                        "periods": {},
                        "resultaat": "",
                        "beoordeling": "",
                        "lijst": "",
                        "pdf_examens": None,
                        "pdf_tekorten": None,
                        "bron": os.path.basename(path),
                    }
                    last_code = None
                    in_courses = False
                    continue

                if cur is None:
                    continue

                mp = PLAN_RE.search(line)
                if mp and not cur["plancode"]:
                    cur["plancode"] = mp.group("plancode").strip()
                    cur["planname"] = mp.group("planname").strip()
                    continue

                if CODE_RE.match(line):
                    mt = TAIL_RE.search(line)
                    if mt:
                        code = line.split()[0]
                        disp, val = grade_to_value(mt.group("grade"))
                        if code in cur["grades"]:
                            problems.append(
                                f"{cur['bron']}: student {cur['nr']} {cur['naam']} "
                                f"heeft vak {code} meer dan één keer; laatste punt gebruikt."
                            )
                        cur["grades"][code] = (disp, val)
                        last_code = code
                        in_courses = True
                        head = line[: mt.start()]
                        mh = HEAD_RE.match(head)
                        if mh:
                            naam = mh.group("name").strip()
                            sp = mh.group("sp").replace(",", ".")
                            lect = mh.group("lect").strip()
                            # Examenperiode (S01, S02, S12, Z02, ...). Wisselt
                            # die tussen juni en september, dan is het vak in
                            # de tweede zit opnieuw afgelegd - ook als het punt
                            # gelijk bleef.
                            cur["periods"][code] = mh.group("period")
                        else:
                            parts = head.split(maxsplit=1)
                            naam = parts[1].strip() if len(parts) > 1 else ""
                            sp, lect = "", ""
                        courses_registry.setdefault(
                            code, {"naam": naam, "sp": sp, "lector": lect}
                        )
                        continue

                if "Totaal aantal SP" in line or "Default SP vereist" in line:
                    in_courses = False
                if "Totaal aantal examens" in line:
                    me = EXAMENS_RE.search(line)
                    if me:
                        cur["pdf_examens"] = int(me.group(1))
                    mb = BEOORDELING_RE.search(line)
                    if mb:
                        cur["beoordeling"] = f"{mb.group(1)} - {mb.group(2).strip()}"
                    in_courses = False
                    continue
                if "Behaald resultaat" in line:
                    mr = RESULTAAT_RE.search(line)
                    if mr:
                        cur["resultaat"] = int(mr.group(1))
                    in_courses = False
                if "Aantal tekorten" in line:
                    mt = TEKORTEN_RE.search(line)
                    if mt:
                        cur["pdf_tekorten"] = int(mt.group(1))
                    ml = LIJST_RE.search(line)
                    if ml:
                        cur["lijst"] = f"{ml.group(1)} - {ml.group(2).capitalize()} lijst"
                    in_courses = False
                    continue
                if not cur["lijst"]:
                    ml = LIJST_RE.search(line)
                    if ml and "tekorten" not in line.lower():
                        cur["lijst"] = f"{ml.group(1)} - {ml.group(2).capitalize()} lijst"

                if in_courses and last_code and re.match(r"^[a-zà-ÿ(]", line.strip()):
                    frag = line.strip()
                    courses_registry[last_code]["naam"] += " " + frag

        close()
    return students


# ---------------------------------------------------------------------------
# 4. Controle per student
# ---------------------------------------------------------------------------
def check_student(s):
    """Geef (ok, lijst_van_meldingen)."""
    msgs = []
    found = len(s["grades"])
    fails = sum(1 for disp, v in s["grades"].values() if is_deficit(disp, v))

    if s["pdf_examens"] is not None and found != s["pdf_examens"]:
        msgs.append(
            f"aantal vakken gevonden ({found}) ≠ 'Totaal aantal examens' in PDF "
            f"({s['pdf_examens']})"
        )
    if s["pdf_tekorten"] is not None and fails != s["pdf_tekorten"]:
        msgs.append(
            f"aantal punten < {PASS_THRESHOLD} berekend ({fails}) ≠ 'Aantal tekorten' "
            f"in PDF ({s['pdf_tekorten']})"
        )
    for code, (disp, val) in s["grades"].items():
        if val is not None and not (0 <= val <= 20):
            msgs.append(f"vak {code}: punt {disp} ligt buiten 0-20")
    if not s["plancode"]:
        msgs.append("geen opleiding/plan gevonden")
    return (len(msgs) == 0), msgs


# ---------------------------------------------------------------------------
# 4a. Deliberatielijsten van de eerste en de tweede zit combineren
#
# De kopregel van elke pagina vermeldt om welke zit het gaat, bv:
#     "DELIBERATIELIJST Deliberatie 2e zit 2025-2026 Sortering: Percentage"
#
# De lijst van de tweede zit bevat ALLE vakken van de student met het op dat
# moment geldende punt, dus ook de vakken die al in de eerste zit geslaagd
# waren. Daarom wordt de kolom "2e" alleen gevuld wanneer het punt effectief
# afwijkt van de eerste zit: dat zijn de herkansingen.
# ---------------------------------------------------------------------------
ZIT2_RE = re.compile(r"\b(2e|2de|tweede)\s*zit\b", re.IGNORECASE)
ZIT1_RE = re.compile(r"\b(1e|1ste|eerste)\s*zit\b", re.IGNORECASE)


def detect_zit(path):
    """Geef 1 of 2 terug: welke examenperiode beschrijft deze deliberatielijst."""
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return 1
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return 1
    if ZIT2_RE.search(text):
        return 2
    if ZIT1_RE.search(text):
        return 1
    return 1          # geen aanduiding: behandel als eerste zit


def merge_zit_students(students, problems):
    """
    Voeg studenten uit de eerste- en tweedezitlijsten samen op studentnummer.

    Geeft een lijst samengevoegde studentdicts terug met per vak
    {"d1","v1","d2","v2"} in de sleutel "grades2".
    """
    per_nr = {}
    for s in students:
        per_nr.setdefault(s["nr"], []).append(s)

    merged = []
    for nr, recs in per_nr.items():
        zit1 = [r for r in recs if r.get("zit", 1) == 1]
        zit2 = [r for r in recs if r.get("zit", 1) == 2]
        if len(zit1) > 1 or len(zit2) > 1:
            problems.append(
                f"Student {nr} {recs[0]['naam']} komt meer dan één keer voor in "
                f"dezelfde zit; het laatst gelezen bestand is gebruikt."
            )
        r1 = zit1[-1] if zit1 else None
        r2 = zit2[-1] if zit2 else None
        laatste = r2 or r1          # meest recente gegevens voor meta/controle

        grades2 = {}
        for code in set(list((r1 or {"grades": {}})["grades"]) +
                        list((r2 or {"grades": {}})["grades"])):
            g1 = r1["grades"].get(code) if r1 else None
            g2 = r2["grades"].get(code) if r2 else None
            if g1 and g2:
                # Het vak telt als herkanst wanneer het punt wijzigde OF de
                # examenperiode wijzigde (bv. S01 in juni -> Z02 in september).
                # Dat tweede is nodig omdat een herkansing met hetzelfde punt
                # anders onzichtbaar zou blijven.
                p1 = (r1.get("periods") or {}).get(code, "")
                p2 = (r2.get("periods") or {}).get(code, "")
                herkanst = (g2[0] != g1[0]) or (p1 and p2 and p1 != p2)
                if herkanst:
                    d1, v1, d2, v2 = g1[0], g1[1], g2[0], g2[1]
                else:
                    d1, v1, d2, v2 = g1[0], g1[1], "", None
            elif g1:
                d1, v1, d2, v2 = g1[0], g1[1], "", None
            else:                    # enkel in de tweedezitlijst
                d1, v1, d2, v2 = "", None, g2[0], g2[1]
            grades2[code] = {"d1": d1, "v1": v1, "d2": d2, "v2": v2}

        zitten = ("1+2" if (r1 and r2) else "1" if r1 else "2")
        m = dict(laatste)
        m["grades2"] = grades2
        m["zitten"] = zitten
        m["bron"] = " + ".join(sorted({r["bron"] for r in recs}))
        merged.append(m)
    return merged


def check_student_merged(s):
    """Controle op de samengevoegde (eind)toestand van een student."""
    msgs = []
    found = len(s["grades2"])
    fails = 0
    for g in s["grades2"].values():
        de, ve = eind_punt(g)
        if is_deficit(de, ve):
            fails += 1
    if s.get("pdf_examens") is not None and found != s["pdf_examens"]:
        msgs.append(
            f"aantal vakken gevonden ({found}) ≠ 'Totaal aantal examens' in PDF "
            f"({s['pdf_examens']})"
        )
    if s.get("pdf_tekorten") is not None and fails != s["pdf_tekorten"]:
        msgs.append(
            f"aantal tekorten na tweede zit ({fails}) ≠ 'Aantal tekorten' in PDF "
            f"({s['pdf_tekorten']})"
        )
    if not s.get("plancode"):
        msgs.append("geen opleiding/plan gevonden")
    return (len(msgs) == 0), msgs


# ---------------------------------------------------------------------------
# 4b. Tweede PDF-type: "Rapport Vaststelling Punten" (niet-diplomajaar)
#
# Dit rapport toont per student ALLE vakken met het resultaat van de eerste
# en (indien van toepassing) de tweede zit:
#
#     Opleidingsonderdeel                              SP  1e zit  2de zit
#     De expert in het evidence based zorgproces ...    5    08      11
#     De professional als beheerder van kwaliteits...   5    16
#
# Er staan in dit rapport geen controlegetallen ("Totaal aantal examens" /
# "Aantal tekorten"), dus de kruiscontrole van de deliberatielijsten is hier
# niet mogelijk.
# ---------------------------------------------------------------------------
VAST_MARKER = "Rapport Vaststelling Punten"

VAST_PERIODE_RE = re.compile(r"^Periode\s+(?P<periode>\d{4})\s*-\s*(?P<soort>.+?)\s*$")
VAST_TRACK_RE = re.compile(r"^(?P<naam>.+?)\s*\((?P<contract>[^()]*[Cc]ontract)\)\s*$")
VAST_STUDENT_RE = re.compile(r"^(?P<naam>.+?)\s*\((?P<nr>\d{8})\)\s*$")
# Vakregel: naam (mogelijk afgekapt) + studiepunten + punt 1e zit [+ punt 2de zit]
VAST_COURSE_RE = re.compile(
    r"^(?P<name>.+?)\s+(?P<sp>\d{1,2})\s+"
    r"(?P<g1>\d{1,2}|[A-Z]{1,4})"
    r"(?:\s+(?P<g2>\d{1,2}|[A-Z]{1,4}))?\s*$"
)
# Vakcode binnen de (vaak afgekapte) vaknaam, bv "(2022GENVEV; S01)"
VAST_CODE_RE = re.compile(r"(\d{4}[A-Z]{2,})\s*;")

# Korte richtingslabels voor dit rapport (de namen wijken licht af van de
# deliberatielijsten, daarom herkenning op trefwoord).
VAST_TRACK_KEYWORDS = [
    ("leiderschap", "LGZ - Leiderschap", "Master - Leiderschap in gezondheid en zorg"),
    ("onderzoeker", "OGZ - Onderzoeker", "Master - Onderzoeker in gezondheid en zorg"),
    ("verpleegkundig spec", "VES - Verpleegk. spec.", "Master - Verpleegkundig specialist"),
    ("vroedvrouw", "VRS - Vroedvrouw spec.", "Master - Vroedvrouw specialist"),
]


def is_vaststelling_pdf(path):
    """True als dit een 'Rapport Vaststelling Punten' is (niet-diplomajaar)."""
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                return False
            text = pdf.pages[0].extract_text() or ""
        return VAST_MARKER.lower() in text.lower()
    except Exception:
        return False


def _clean_course_name(raw):
    """Haal een afgekapt staartje zoals '(2022GENVEV; S01)' of '(50' weg."""
    naam = re.sub(r"\s*\([^)]*\)\s*$", "", raw).strip()   # volledige haakjes
    naam = re.sub(r"\s*\([^)]*$", "", naam).strip()       # afgekapte haakjes
    return naam or raw.strip()


def match_course_key(raw_name, courses_registry):
    """
    Bepaal onder welke sleutel dit vak in de Excel komt.

    Geeft (sleutel, weergavenaam) terug. Wanneer het vak ook in de
    deliberatielijsten voorkomt, worden dezelfde vakcode en dezelfde
    (volledige) vaknaam gebruikt, zodat de tabbladen op elkaar aansluiten.
    """
    kort = _clean_course_name(raw_name)

    # 1. Vakcode staat in de regel en is bekend uit de deliberatielijsten.
    m = VAST_CODE_RE.search(raw_name)
    if m:
        code = m.group(1)
        if code in courses_registry:
            return code, courses_registry[code]["naam"] or kort
        return code, kort

    # 2. Geen code: zoek een uniek vak in de registry dat met deze
    #    (afgekapte) naam begint.
    if kort:
        laag = kort.lower()
        treffers = [c for c, reg in courses_registry.items()
                    if (reg.get("naam") or "").lower().startswith(laag)]
        if len(treffers) == 1:
            code = treffers[0]
            return code, courses_registry[code]["naam"] or kort

    # 3. Onbekend vak: eigen sleutel op basis van de naam.
    return "~" + kort.lower(), kort


GRADE_TOKEN_RE = re.compile(r"^(\d{1,2}|[A-Z]{1,4})$")


def _regels_met_posities(page):
    """
    Geef per tekstregel de woorden mét hun x-positie terug.

    Nodig omdat de kolommen '1e zit' en '2de zit' in de platte tekst niet te
    onderscheiden zijn: staat enkel de tweede zit ingevuld, dan ziet een regel
    er identiek uit als eentje met enkel een eerste zit. De x-positie van het
    punt verraadt wel in welke kolom het staat.
    """
    woorden = page.extract_words() or []
    regels = {}
    for w in woorden:
        regels.setdefault(round(w["top"]), []).append(w)
    uit = []
    for top in sorted(regels):
        ws = sorted(regels[top], key=lambda w: w["x0"])
        uit.append((" ".join(w["text"] for w in ws), ws))
    return uit


def _kolomgrenzen(page):
    """
    Zoek op deze pagina de x-posities van de kolommen 'SP', '1e zit' en
    '2de zit'. Geeft (sp_x, grens_1e, grens_tussen) terug, of None.
    """
    for _, ws in _regels_met_posities(page):
        tekst = " ".join(w["text"] for w in ws)
        if "1e zit" not in tekst or "2de zit" not in tekst:
            continue
        sp = next((w for w in ws if w["text"] == "SP"), None)
        een = next((w for w in ws if w["text"] == "1e"), None)
        twee = next((w for w in ws if w["text"] == "2de"), None)
        if een and twee:
            sp_x = sp["x0"] if sp else een["x0"] - 55
            return sp_x, een["x0"] - 15, (een["x1"] + twee["x0"]) / 2
    return None


def _split_vakregel(ws, sp_x, grens_1e, grens_tussen):
    """
    Verdeel de woorden van een vakregel over naam, studiepunten en de punten
    van de eerste en de tweede zit, op basis van hun x-positie.

    Geeft (naam, sp, punt1, punt2) terug, of None als dit geen vakregel is.
    """
    naam_w, sp_w, g1, g2 = [], None, "", ""
    for w in ws:
        t = w["text"]
        if w["x0"] >= grens_1e:                       # puntenkolommen
            if not GRADE_TOKEN_RE.match(t):
                return None
            if w["x0"] < grens_tussen:
                g1 = t if not g1 else g1
            else:
                g2 = t if not g2 else g2
        elif sp_w is None and w["x0"] >= sp_x - 10 and GRADE_TOKEN_RE.match(t):
            sp_w = t                                  # studiepuntenkolom
        else:
            naam_w.append(t)
    if not naam_w or (not g1 and not g2):
        return None
    return " ".join(naam_w), (sp_w or ""), g1, g2


def parse_vaststelling_pdf(path, courses_registry, problems):
    """Lees één 'Vaststelling Punten'-PDF. Geeft een lijst studentdicts terug."""
    students = []
    cur = None
    jaar = "Onbekend"
    track_naam = ""
    track_kort = ""

    with pdfplumber.open(path) as pdf:
        # De kolomposities zijn in het hele rapport dezelfde. Loopt de
        # vakkenlijst van een student door op een volgende pagina, dan staat
        # daar geen kopregel meer; daarom onthouden we de laatst gevonden
        # posities en gebruiken we die verder.
        laatste_grenzen = None
        for page in pdf.pages:
            if laatste_grenzen is None:
                for p in pdf.pages:
                    laatste_grenzen = _kolomgrenzen(p)
                    if laatste_grenzen:
                        break
            grenzen = _kolomgrenzen(page) or laatste_grenzen
            if grenzen:
                laatste_grenzen = grenzen
            for line, ws in _regels_met_posities(page):
                line = line.strip()
                if not line or "Page " in line or line.startswith("Opleidingsonderdeel"):
                    continue
                if VAST_MARKER.lower() in line.lower():
                    continue

                mp = VAST_PERIODE_RE.match(line)
                if mp:
                    soort = mp.group("soort").lower()
                    if "schakel" in soort or "voorbereiding" in soort:
                        jaar = "Schakeljaar"
                    elif "master" in soort:
                        jaar = "Masterjaar"
                    continue

                ms = VAST_STUDENT_RE.match(line)
                if ms:
                    cur = {
                        "nr": ms.group("nr"),
                        "naam": ms.group("naam").strip(),
                        "status_jaar": jaar,
                        "track": track_naam or jaar,
                        "track_short": track_kort or jaar,
                        "grades": {},        # sleutel -> (punt1, punt2)
                        "bron": os.path.basename(path),
                    }
                    students.append(cur)
                    continue

                mt = VAST_TRACK_RE.match(line)
                if mt:
                    ruw = mt.group("naam").strip()
                    track_naam, track_kort = ruw, ruw
                    for sleutel, kort, vol in VAST_TRACK_KEYWORDS:
                        if sleutel in ruw.lower():
                            track_naam, track_kort = vol, kort
                            break
                    cur = None
                    continue

                # Vakregel: bij voorkeur op x-positie ontleden (dan weten we
                # zeker bij welke zit een punt hoort), anders op tekstpatroon.
                velden = None
                if grenzen and cur is not None:
                    velden = _split_vakregel(ws, *grenzen)
                if velden is None and cur is not None:
                    mc = VAST_COURSE_RE.match(line)
                    if mc:
                        velden = (mc.group("name"), mc.group("sp"),
                                  mc.group("g1"), mc.group("g2") or "")
                        # Terugval op tekstherkenning: dan is niet met zekerheid
                        # te zeggen bij welke zit een enkel punt hoort.
                        if not mc.group("g2"):
                            problems.append(
                                f"{os.path.basename(path)}: bij student "
                                f"{cur['nr']} {cur['naam']} kon de kolom van "
                                f"'{line[:45]}...' niet op positie bepaald "
                                f"worden; punt toegekend aan de eerste zit. "
                                f"CONTROLEER dit vak."
                            )
                if velden is not None:
                    ruwe_naam, sp_txt, g1_txt, g2_txt = velden
                    key, naam = match_course_key(ruwe_naam, courses_registry)
                    d1, v1 = grade_to_value(g1_txt) if g1_txt else ("", None)
                    d2, v2 = grade_to_value(g2_txt) if g2_txt else ("", None)
                    if key in cur["grades"]:
                        problems.append(
                            f"{cur['bron']}: student {cur['nr']} {cur['naam']} heeft "
                            f"vak '{naam}' meer dan één keer; laatste punt gebruikt."
                        )
                    cur["grades"][key] = {
                        "naam": naam, "sp": sp_txt,
                        "d1": d1, "v1": v1, "d2": d2, "v2": v2,
                    }
                    continue

    return students


def eind_punt(g):
    """Geef (weergave, waarde) van het eindresultaat: 2de zit indien aanwezig."""
    if g["d2"]:
        return g["d2"], g["v2"]
    return g["d1"], g["v1"]


def consolidate_zit2_keys(students):
    """
    Voeg vakken samen die door afkapping onder twee sleutels terechtkwamen.

    Bij de ene student staat de vakcode nog in de regel ("2022GENVEV"), bij de
    andere is die weggevallen door de kolombreedte. Zonder deze stap zou
    hetzelfde vak twee kolommen krijgen.
    """
    met_code = {}   # sleutel met echte code -> weergavenaam (kleine letters)
    for s in students:
        for key, g in s["grades"].items():
            if not key.startswith("~"):
                met_code.setdefault(key, (g["naam"] or "").lower())

    hermap = {}
    for s in students:
        for key, g in s["grades"].items():
            if not key.startswith("~") or key in hermap:
                continue
            kort = key[1:]
            treffers = [c for c, naam in met_code.items() if naam.startswith(kort)]
            if len(treffers) == 1:
                hermap[key] = treffers[0]

    if not hermap:
        return 0
    for s in students:
        for oud, nieuw in hermap.items():
            if oud in s["grades"]:
                g = s["grades"].pop(oud)
                s["grades"].setdefault(nieuw, g)
    return len(hermap)


# ---------------------------------------------------------------------------
# 5. Excel opbouwen
# ---------------------------------------------------------------------------
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006", bold=True)
CREDIT_FILL = PatternFill("solid", fgColor="E2EFDA")
HEAD_FILL = PatternFill("solid", fgColor="305496")
HEAD_FONT = Font(color="FFFFFF", bold=True)
META_FILL = PatternFill("solid", fgColor="D9E1F2")
BAD_FILL = PatternFill("solid", fgColor="FFEB9C")
GOOD_FONT = Font(color="006100")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
VERT = Alignment(textRotation=90, vertical="bottom", horizontal="center", wrap_text=True)


def kop_vak(code, naam):
    """Kolomkop van een vak: altijd met de vakcode erbij."""
    naam = (naam or "").strip()
    if not code or code.startswith("~"):
        return naam
    return f"{code} - {naam}" if naam else code


def sorteer_studenten(students, op_resultaat=True):
    """
    Sorteer van hoog naar laag.

    Op de deliberatietabbladen gebeurt dat op 'Resultaat %'; op de tabbladen
    van het niet-diplomajaar (waar de PDF geen percentage vermeldt) op het
    gemiddelde van de eindpunten. Bij een gelijke stand telt de naam.
    """
    def sleutel(s):
        waarde = None
        if op_resultaat:
            r = s.get("resultaat", "")
            waarde = r if isinstance(r, (int, float)) else None
        if waarde is None:
            punten = []
            for g in (s.get("grades2") or s.get("grades") or {}).values():
                if isinstance(g, dict):
                    _, v = eind_punt(g)
                else:
                    v = g[1]
                if v is not None:
                    punten.append(v)
            waarde = sum(punten) / len(punten) if punten else -1
        return (-waarde, s["naam"].lower())
    return sorted(students, key=sleutel)


def zet_filter(ws, laatste_kolom, laatste_rij):
    """Zet de filterknoppen aan op de kopregel (rij 2) van een matrixtabblad."""
    if laatste_rij < 3:
        return
    ws.auto_filter.ref = f"A2:{get_column_letter(laatste_kolom)}{laatste_rij}"


def build_matrix_sheet(ws, students, courses_registry, show_track=False):
    META = ["Studentnr", "Naam", "Jaar"]
    if show_track:
        META.append("Richting")
    META += ["Lijst", "Resultaat %", "Eindbeoordeling"]
    nmeta = len(META)

    codes = sorted(
        {c for s in students for c in s["grades"]},
        key=lambda c: (int(re.match(r"\d+", c).group()), c),
    )

    # Kopregel: het label staat op rij 2 (de filterregel); rij 1 blijft vrij
    # voor de gedraaide vaknamen.
    for j, label in enumerate(META, start=1):
        c1 = ws.cell(row=1, column=j, value="")
        c1.fill, c1.border = HEAD_FILL, BORDER
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER

    for k, code in enumerate(codes):
        col = nmeta + 1 + k
        reg = courses_registry.get(code, {"naam": "", "sp": "", "lector": ""})
        top = ws.cell(row=1, column=col, value=kop_vak(code, reg["naam"]))
        top.fill, top.font, top.alignment, top.border = HEAD_FILL, HEAD_FONT, VERT, BORDER
        bot = ws.cell(row=2, column=col, value=code)
        bot.fill, bot.font, bot.alignment, bot.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
        comment_txt = f"{code}\n{reg['naam']}\nStudiepunten: {reg['sp']}\nLector: {reg['lector']}"
        bot.comment = Comment(comment_txt, "puntenlijst")
        ws.column_dimensions[get_column_letter(col)].width = 5.5

    extra = ["Gem.", "# Tekorten", "Examens (PDF)", "Tekorten (PDF)", "Controle"]
    for e, label in enumerate(extra):
        col = nmeta + 1 + len(codes) + e
        c1 = ws.cell(row=1, column=col, value="")
        c1.fill, c1.border = HEAD_FILL, BORDER
        c = ws.cell(row=2, column=col, value=label)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
        ws.column_dimensions[get_column_letter(col)].width = 13

    ws.row_dimensions[1].height = 150

    for i, s in enumerate(sorteer_studenten(students)):
        r = 3 + i
        meta_vals = [s["nr"], s["naam"], s["status_jaar"]]
        if show_track:
            meta_vals.append(s.get("track_short", s.get("track", "")))
        meta_vals += [s["lijst"], s["resultaat"], s["beoordeling"]]
        for j, v in enumerate(meta_vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.fill, c.border = META_FILL, BORDER
            if META[j - 1] == "Resultaat %":
                c.alignment = CENTER

        numeric_vals = []
        fails = 0
        for k, code in enumerate(codes):
            col = nmeta + 1 + k
            cell = ws.cell(row=r, column=col)
            cell.alignment, cell.border = CENTER, BORDER
            if code in s["grades"]:
                disp, val = s["grades"][code]
                cell.value = val if val is not None else disp
                if val is not None:
                    numeric_vals.append(val)
                if is_deficit(disp, val):
                    fails += 1
                    cell.fill, cell.font = RED_FILL, RED_FONT
                elif is_credit(disp, val):
                    cell.fill = CREDIT_FILL
            else:
                cell.value = ""

        ok, _ = check_student(s)
        gem = round(sum(numeric_vals) / len(numeric_vals), 1) if numeric_vals else ""
        ctrl_vals = [
            gem, fails, s["pdf_examens"], s["pdf_tekorten"],
            "OK" if ok else "CONTROLEER",
        ]
        for e, v in enumerate(ctrl_vals):
            col = nmeta + 1 + len(codes) + e
            c = ws.cell(row=r, column=col, value=v)
            c.alignment, c.border = CENTER, BORDER
            if e == 4:
                if ok:
                    c.font = GOOD_FONT
                else:
                    c.fill, c.font = BAD_FILL, Font(bold=True, color="9C6500")

    META_WIDTHS = {
        "Studentnr": 12, "Naam": 26, "Jaar": 12, "Richting": 22,
        "Lijst": 16, "Resultaat %": 11, "Eindbeoordeling": 26,
    }
    for j, label in enumerate(META, start=1):
        ws.column_dimensions[get_column_letter(j)].width = META_WIDTHS.get(label, 14)

    ws.freeze_panes = ws.cell(row=3, column=nmeta + 1)
    zet_filter(ws, nmeta + len(codes) + len(extra), 2 + len(students))


def build_matrix_sheet_dual(ws, students, courses_registry, show_track=False):
    """
    Deliberatietabblad met eerste EN tweede zit: per vak drie kolommen
    (1e zit / 2de zit / eind). Wordt gebruikt zodra er ook een
    tweedezitlijst is aangeleverd.
    """
    META = ["Studentnr", "Naam", "Jaar"]
    if show_track:
        META.append("Richting")
    META += ["Zit", "Lijst", "Resultaat %", "Eindbeoordeling"]
    nmeta = len(META)

    codes = sorted(
        {c for s in students for c in s["grades2"]},
        key=lambda c: (int(re.match(r"\d+", c).group()), c),
    )

    for j, label in enumerate(META, start=1):
        c1 = ws.cell(row=1, column=j, value="")
        c1.fill, c1.border = HEAD_FILL, BORDER
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER

    for k, code in enumerate(codes):
        col = nmeta + 1 + k * 3
        reg = courses_registry.get(code, {"naam": "", "sp": "", "lector": ""})
        top = ws.cell(row=1, column=col, value=kop_vak(code, reg["naam"]))
        top.fill, top.font, top.alignment, top.border = HEAD_FILL, HEAD_FONT, VERT, BORDER
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        for e, sub in enumerate(["1e", "2e", "eind"]):
            c = ws.cell(row=2, column=col + e, value=sub)
            c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
            c.comment = Comment(
                f"{code}\n{reg['naam']}\nStudiepunten: {reg['sp']}\n"
                f"Lector: {reg['lector']}", "puntenlijst")
            ws.column_dimensions[get_column_letter(col + e)].width = 5.5
        # 1e en 2e zit inklapbaar maken; 'eind' blijft altijd zichtbaar.
        groepeer_zitkolommen(ws, col)

    extra = ["Gem. eind", "# Tekorten", "# Herkansingen",
             "Examens (PDF)", "Tekorten (PDF)", "Controle"]
    for e, label in enumerate(extra):
        col = nmeta + 1 + len(codes) * 3 + e
        c1 = ws.cell(row=1, column=col, value="")
        c1.fill, c1.border = HEAD_FILL, BORDER
        c = ws.cell(row=2, column=col, value=label)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
        ws.column_dimensions[get_column_letter(col)].width = 13

    ws.row_dimensions[1].height = 150

    for i, s in enumerate(sorteer_studenten(students)):
        r = 3 + i
        meta_vals = [s["nr"], s["naam"], s["status_jaar"]]
        if show_track:
            meta_vals.append(s.get("track_short", s.get("track", "")))
        meta_vals += [s.get("zitten", ""), s.get("lijst", ""),
                      s.get("resultaat", ""), s.get("beoordeling", "")]
        for j, v in enumerate(meta_vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.fill, c.border = META_FILL, BORDER
            if META[j - 1] in ("Resultaat %", "Zit"):
                c.alignment = CENTER

        numeriek, tekorten, herkansingen = [], 0, 0
        for k, code in enumerate(codes):
            col = nmeta + 1 + k * 3
            for e in range(3):
                cell = ws.cell(row=r, column=col + e)
                cell.alignment, cell.border = CENTER, BORDER
                if e == 1:
                    cell.fill = ZIT2_FILL
                elif e == 2:
                    cell.fill = EIND_FILL
            g = s["grades2"].get(code)
            if not g:
                continue
            de, ve = eind_punt(g)
            if g["d2"]:
                herkansingen += 1
            if ve is not None:
                numeriek.append(ve)
            if is_deficit(de, ve):
                tekorten += 1
            for e, (disp, val) in enumerate([(g["d1"], g["v1"]),
                                             (g["d2"], g["v2"]), (de, ve)]):
                cell = ws.cell(row=r, column=col + e)
                if not disp:
                    continue
                cell.value = val if val is not None else disp
                if is_deficit(disp, val):
                    cell.fill, cell.font = RED_FILL, RED_FONT
                elif is_credit(disp, val):
                    cell.fill = CREDIT_FILL

        ok, _ = check_student_merged(s)
        gem = round(sum(numeriek) / len(numeriek), 1) if numeriek else ""
        ctrl = [gem, tekorten, herkansingen, s.get("pdf_examens"),
                s.get("pdf_tekorten"), "OK" if ok else "CONTROLEER"]
        for e, v in enumerate(ctrl):
            c = ws.cell(row=r, column=nmeta + 1 + len(codes) * 3 + e, value=v)
            c.alignment, c.border = CENTER, BORDER
            if e == 5:
                if ok:
                    c.font = GOOD_FONT
                else:
                    c.fill, c.font = BAD_FILL, Font(bold=True, color="9C6500")

    META_WIDTHS = {"Studentnr": 12, "Naam": 26, "Jaar": 12, "Richting": 22,
                   "Zit": 7, "Lijst": 16, "Resultaat %": 11,
                   "Eindbeoordeling": 26}
    for j, label in enumerate(META, start=1):
        ws.column_dimensions[get_column_letter(j)].width = META_WIDTHS.get(label, 14)

    ws.freeze_panes = ws.cell(row=3, column=nmeta + 1)
    zet_filter(ws, nmeta + len(codes) * 3 + len(extra), 2 + len(students))


def build_legend_sheet(ws, courses_registry):
    headers = ["Vakcode", "Vaknaam", "Studiepunten", "Lector(en)"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BORDER
    for i, code in enumerate(sorted(courses_registry), start=2):
        reg = courses_registry[code]
        ws.cell(row=i, column=1, value=code).border = BORDER
        ws.cell(row=i, column=2, value=reg["naam"]).border = BORDER
        ws.cell(row=i, column=3, value=reg["sp"]).border = BORDER
        ws.cell(row=i, column=4, value=reg["lector"]).border = BORDER
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 40
    ws.freeze_panes = "A2"
    if courses_registry:
        ws.auto_filter.ref = f"A1:D{len(courses_registry) + 1}"


def build_check_sheet(ws, all_students):
    headers = ["Bron (PDF)", "Studentnr", "Naam", "Jaar", "Opleiding",
               "Status controle", "Detail"]
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BORDER
    r = 2
    n_ok = n_bad = 0
    for s in sorted(all_students, key=lambda x: (x["track"], x["naam"].lower())):
        ok, msgs = (check_student_merged(s) if "grades2" in s
                    else check_student(s))
        n_ok += ok
        n_bad += (not ok)
        ws.cell(row=r, column=1, value=s["bron"])
        ws.cell(row=r, column=2, value=s["nr"])
        ws.cell(row=r, column=3, value=s["naam"])
        ws.cell(row=r, column=4, value=s["status_jaar"])
        ws.cell(row=r, column=5, value=s["track"])
        status_cell = ws.cell(row=r, column=6, value="OK" if ok else "CONTROLEER")
        ws.cell(row=r, column=7, value="; ".join(msgs))
        if ok:
            status_cell.font = GOOD_FONT
        else:
            status_cell.fill, status_cell.font = BAD_FILL, Font(bold=True, color="9C6500")
        r += 1
    for col, w in zip("ABCDEFG", [28, 12, 26, 12, 42, 16, 70]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    if r > 2:
        ws.auto_filter.ref = f"A1:G{r - 1}"
    return n_ok, n_bad


ZIT2_FILL = PatternFill("solid", fgColor="FFF2CC")   # kolom '2e zit'
EIND_FILL = PatternFill("solid", fgColor="EDEDED")   # kolom 'eind'


def groepeer_zitkolommen(ws, col):
    """
    Maak de kolommen '1e' en '2e' van één vak inklapbaar.

    Het bestand opent met deze kolommen INGEKLAPT, zodat je meteen het
    overzicht van de eindresultaten ziet. Met de plusjes boven de kolommen
    open je één vak; met het knopje '2' linksboven klap je alles in één keer
    open (en met '1' weer dicht). De kolom 'eind' blijft altijd zichtbaar.
    """
    try:
        ws.sheet_properties.outlinePr.summaryRight = True
        ws.column_dimensions.group(
            get_column_letter(col), get_column_letter(col + 1),
            outline_level=1, hidden=True,
        )
    except Exception:
        pass


def build_zit2_sheet(ws, students, show_track=False):
    """
    Matrixtabblad voor het niet-diplomajaar: per vak drie kolommen
    (1e zit / 2de zit / eindresultaat).
    """
    META = ["Studentnr", "Naam", "Jaar"]
    if show_track:
        META.append("Richting")
    nmeta = len(META)

    # Alle vakken die in dit tabblad voorkomen, met hun weergavenaam.
    namen = {}
    for s in students:
        for key, g in s["grades"].items():
            namen.setdefault(key, g["naam"])
    keys = sorted(namen, key=lambda k: (namen[k] or "").lower())

    # ---- koprijen ----
    for j, label in enumerate(META, start=1):
        c1 = ws.cell(row=1, column=j, value="")
        c1.fill, c1.border = HEAD_FILL, BORDER
        c = ws.cell(row=2, column=j, value=label)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER

    for k, key in enumerate(keys):
        col = nmeta + 1 + k * 3
        top = ws.cell(row=1, column=col, value=kop_vak(key, namen[key]))
        top.fill, top.font, top.alignment, top.border = HEAD_FILL, HEAD_FONT, VERT, BORDER
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        for e, sub in enumerate(["1e", "2e", "eind"]):
            c = ws.cell(row=2, column=col + e, value=sub)
            c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
            comment_txt = namen[key] + (f"\nVakcode: {key}" if not key.startswith("~") else "")
            c.comment = Comment(comment_txt, "puntenlijst")
            ws.column_dimensions[get_column_letter(col + e)].width = 5.5
        groepeer_zitkolommen(ws, col)

    extra = ["Gem. eind", "# Tekorten", "# Vakken", "# Tweede zit"]
    for e, label in enumerate(extra):
        col = nmeta + 1 + len(keys) * 3 + e
        c1 = ws.cell(row=1, column=col, value="")
        c1.fill, c1.border = HEAD_FILL, BORDER
        c = ws.cell(row=2, column=col, value=label)
        c.fill, c.font, c.alignment, c.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
        ws.column_dimensions[get_column_letter(col)].width = 12

    ws.row_dimensions[1].height = 150

    # ---- studentrijen ----
    for i, s in enumerate(sorteer_studenten(students, op_resultaat=False)):
        r = 3 + i
        meta_vals = [s["nr"], s["naam"], s["status_jaar"]]
        if show_track:
            meta_vals.append(s.get("track_short", ""))
        for j, v in enumerate(meta_vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.fill, c.border = META_FILL, BORDER

        numeriek, tekorten, n_tweede = [], 0, 0
        for k, key in enumerate(keys):
            col = nmeta + 1 + k * 3
            g = s["grades"].get(key)
            for e in range(3):
                cell = ws.cell(row=r, column=col + e)
                cell.alignment, cell.border = CENTER, BORDER
                if e == 1:
                    cell.fill = ZIT2_FILL
                elif e == 2:
                    cell.fill = EIND_FILL
            if not g:
                continue
            de, ve = eind_punt(g)
            if g["d2"]:
                n_tweede += 1
            if ve is not None:
                numeriek.append(ve)
            if is_deficit(de, ve):
                tekorten += 1

            paren = [(g["d1"], g["v1"]), (g["d2"], g["v2"]), (de, ve)]
            for e, (disp, val) in enumerate(paren):
                cell = ws.cell(row=r, column=col + e)
                if not disp:
                    continue
                cell.value = val if val is not None else disp
                if is_deficit(disp, val):
                    cell.fill, cell.font = RED_FILL, RED_FONT
                elif is_credit(disp, val):
                    cell.fill = CREDIT_FILL

        gem = round(sum(numeriek) / len(numeriek), 1) if numeriek else ""
        for e, v in enumerate([gem, tekorten, len(s["grades"]), n_tweede]):
            c = ws.cell(row=r, column=nmeta + 1 + len(keys) * 3 + e, value=v)
            c.alignment, c.border = CENTER, BORDER

    META_WIDTHS = {"Studentnr": 12, "Naam": 26, "Jaar": 12, "Richting": 22}
    for j, label in enumerate(META, start=1):
        ws.column_dimensions[get_column_letter(j)].width = META_WIDTHS.get(label, 14)

    ws.freeze_panes = ws.cell(row=3, column=nmeta + 1)
    zet_filter(ws, nmeta + len(keys) * 3 + len(extra), 2 + len(students))


def sheet_title(name, used):
    """Maak een geldige, unieke Excel-tabbladnaam (max 31 tekens)."""
    bad = r'[]:*?/\\'
    clean = "".join("-" if ch in bad else ch for ch in name)[:31]
    base = clean
    i = 2
    while clean.lower() in used:
        suffix = f" ({i})"
        clean = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(clean.lower())
    return clean


# ---------------------------------------------------------------------------
# 6. Hoofdfunctie: van PDF-paden naar Excel-bestand
# ---------------------------------------------------------------------------
def generate(pdf_paths, out_dir=None, log=print):
    """
    Verwerk een lijst PDF-bestanden en schrijf het Excel-bestand.

    pdf_paths : lijst van paden naar PDF-bestanden
    out_dir   : map waar het Excel-bestand komt (standaard: map van 1e PDF)
    log       : functie die voortgangsberichten ontvangt (bv. print)

    Geeft een dict terug met: out_path, n_students, n_ok, n_bad,
    groups (dict richting -> aantal), problems (lijst meldingen).
    """
    pdfs = sorted(pdf_paths)
    if not pdfs:
        raise ValueError("Geen PDF-bestanden opgegeven.")

    courses_registry = {}
    problems = []
    all_students = []
    zit2_students = []

    # Eerst de deliberatielijsten: die vullen de vakkenregistratie met de
    # volledige vaknamen. De 'Vaststelling Punten'-rapporten kunnen daar
    # daarna op aansluiten (zelfde vaknamen en vakcodes).
    delib_pdfs, vast_pdfs = [], []
    for p in pdfs:
        (vast_pdfs if is_vaststelling_pdf(p) else delib_pdfs).append(p)

    zitten_gezien = set()
    for p in delib_pdfs:
        zit = detect_zit(p)
        zitten_gezien.add(zit)
        log(f"Inlezen (deliberatie, {zit}e zit): {os.path.basename(p)}")
        try:
            students = parse_pdf(p, courses_registry, problems)
        except Exception as e:
            problems.append(f"Kon {os.path.basename(p)} niet verwerken: {e}")
            log(f"   !! overgeslagen door fout: {e}")
            continue
        for s in students:
            jaar, track, track_short = classify_track(
                s["plancode"], s["planname"], s["status"]
            )
            s["status_jaar"] = jaar
            s["track"] = track
            s["track_short"] = track_short
            s["zit"] = zit
        all_students.extend(students)
        log(f"   {len(students)} studenten gevonden")

    # Alleen wanneer BEIDE zittijden zijn aangeleverd, worden de studenten op
    # studentnummer samengevoegd tot één rij met kolommen 1e zit / 2de zit /
    # eind. Levert iemand enkel de tweedezitlijst aan, dan is die lijst al
    # volledig en blijft de gewone (smalle) opmaak behouden.
    dual = {1, 2} <= zitten_gezien and len(all_students) > 0
    if dual:
        voor = len(all_students)
        all_students = merge_zit_students(all_students, problems)
        log("")
        log(f"Eerste en tweede zit samengevoegd: {voor} inschrijvingen -> "
            f"{len(all_students)} studenten")

    for p in vast_pdfs:
        log(f"Inlezen (2de zit / niet-diplomajaar): {os.path.basename(p)}")
        try:
            students = parse_vaststelling_pdf(p, courses_registry, problems)
        except Exception as e:
            problems.append(f"Kon {os.path.basename(p)} niet verwerken: {e}")
            log(f"   !! overgeslagen door fout: {e}")
            continue
        zit2_students.extend(students)
        log(f"   {len(students)} studenten gevonden")

    if zit2_students:
        consolidate_zit2_keys(zit2_students)

    if not all_students and not zit2_students:
        raise ValueError(
            "Geen studenten gevonden in de PDF's. Zijn dit wel deliberatie-PDF's "
            "of 'Rapport Vaststelling Punten'-PDF's?"
        )

    groups = {}
    for s in all_students:
        groups.setdefault(s["track"], []).append(s)

    wb = Workbook()
    wb.remove(wb.active)
    used_titles = set()

    def group_sort_key(track):
        return (0 if track.startswith("Master") else 1 if track.startswith("Schakel") else 2, track)

    def bouw(ws, lijst, registry, show_track=False):
        """
        Kies per tabblad de juiste opmaak: drie kolommen per vak wanneer er
        voor deze groep effectief een eerste én een tweede zit is, anders de
        gewone smalle opmaak.
        """
        heeft_beide = dual and any("1" in s.get("zitten", "") for s in lijst) \
            and any("2" in s.get("zitten", "") for s in lijst)
        if heeft_beide:
            build_matrix_sheet_dual(ws, lijst, registry, show_track=show_track)
        else:
            build_matrix_sheet(ws, lijst, registry, show_track=show_track)

    master_students = [s for s in all_students if s["status_jaar"] == "Masterjaar"]
    if master_students:
        ws = wb.create_sheet(sheet_title("Masterjaar - alle richtingen", used_titles))
        bouw(ws, master_students, courses_registry, show_track=True)

    for track in sorted(groups, key=group_sort_key):
        ws = wb.create_sheet(sheet_title(track, used_titles))
        bouw(ws, groups[track], courses_registry)

    # Extra tabbladen voor het niet-diplomajaar (1e zit / 2de zit / eind)
    zit2_master = [s for s in zit2_students if s["status_jaar"] == "Masterjaar"]
    zit2_schakel = [s for s in zit2_students if s["status_jaar"] == "Schakeljaar"]
    zit2_rest = [s for s in zit2_students
                 if s["status_jaar"] not in ("Masterjaar", "Schakeljaar")]
    if zit2_master:
        ws = wb.create_sheet(sheet_title("Master niet-diplomajaar", used_titles))
        build_zit2_sheet(ws, zit2_master, show_track=True)
    if zit2_schakel:
        ws = wb.create_sheet(sheet_title("Schakeljaar niet-diplomajaar", used_titles))
        build_zit2_sheet(ws, zit2_schakel)
    if zit2_rest:
        ws = wb.create_sheet(sheet_title("Overige niet-diplomajaar", used_titles))
        build_zit2_sheet(ws, zit2_rest, show_track=True)

    if all_students:
        build_check_sheet(wb.create_sheet(sheet_title("Controle", used_titles)),
                          all_students)
    build_legend_sheet(wb.create_sheet(sheet_title("Legende", used_titles)), courses_registry)

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out_name = f"Puntenlijst_{stamp}.xlsx"
    if out_dir is None:
        out_dir = os.path.dirname(pdfs[0])
    if not os.access(out_dir, os.W_OK):
        out_dir = os.path.expanduser("~/Desktop")
        if not os.path.isdir(out_dir) or not os.access(out_dir, os.W_OK):
            out_dir = os.getcwd()
    out_path = os.path.join(out_dir, out_name)
    wb.save(out_path)

    n_ok = sum(1 for s in all_students
               if (check_student_merged(s) if "grades2" in s
                   else check_student(s))[0])
    n_bad = len(all_students) - n_ok

    log("")
    log("=" * 50)
    log("KLAAR")
    if all_students:
        log(f"Deliberatie - studenten : {len(all_students)}")
        log(f"Opleidingen (tabbladen) : {len(groups)}")
        for track in sorted(groups, key=group_sort_key):
            log(f"    - {track}: {len(groups[track])} studenten")
        log(f"Controle OK         : {n_ok}")
        log(f"Controle te checken : {n_bad}")
    if zit2_students:
        log(f"2de zit - studenten     : {len(zit2_students)}")
        if zit2_master:
            log(f"    - Master: {len(zit2_master)} studenten")
        if zit2_schakel:
            log(f"    - Schakeljaar: {len(zit2_schakel)} studenten")
    if problems:
        log("")
        log("Meldingen tijdens het inlezen:")
        for pmsg in problems:
            log(f"    * {pmsg}")
    if n_bad:
        log("")
        log(">>> Bekijk het tabblad 'Controle' in het Excel-bestand.")
    log("")
    log(f"Bestand opgeslagen: {out_path}")

    return {
        "out_path": out_path,
        "n_students": len(all_students) + len(zit2_students),
        "n_delib": len(all_students),
        "n_zit2": len(zit2_students),
        "n_ok": n_ok,
        "n_bad": n_bad,
        "groups": {t: len(g) for t, g in groups.items()},
        "problems": problems,
    }


def collect_pdfs(paths):
    """Zet een mix van mappen en bestanden om in een lijst PDF-paden."""
    pdfs = []
    for p in paths:
        p = os.path.abspath(os.path.expanduser(p))
        if os.path.isdir(p):
            pdfs.extend(glob.glob(os.path.join(p, "*.pdf")))
            pdfs.extend(glob.glob(os.path.join(p, "*.PDF")))
        elif os.path.isfile(p) and p.lower().endswith(".pdf"):
            pdfs.append(p)
    # dubbels verwijderen, volgorde behouden
    seen = set()
    unique = []
    for p in pdfs:
        if p not in seen:
            seen.add(p)
            unique.append(p)
    return unique


# ---------------------------------------------------------------------------
# 7. Commandoregel-gebruik (zoals het oude script)
# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) > 1:
        folder = os.path.abspath(os.path.expanduser(sys.argv[1]))
    else:
        folder = os.path.dirname(os.path.abspath(__file__))

    if not os.path.isdir(folder):
        print(f"FOUT: map bestaat niet: {folder}")
        sys.exit(1)

    pdfs = collect_pdfs([folder])
    if not pdfs:
        print(f"Geen PDF-bestanden gevonden in: {folder}")
        sys.exit(1)

    generate(pdfs, out_dir=folder)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        print("\nER GING IETS MIS, maar het script is netjes gestopt.")
        print(f"Foutmelding: {e}")
        sys.exit(1)
