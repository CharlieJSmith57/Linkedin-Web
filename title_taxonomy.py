"""
title_taxonomy.py — Granular job title classification

Replaces the coarse SENIORITY_RULES in the old enricher.
Returns TWO values for every raw title:
  - discipline  : what kind of work they do (e.g. "Electrical — Protection & Relay")
  - seniority   : their career level    (e.g. "Senior")

Both are stored in the Excel sheet and exported to network.json.
The front-end uses discipline for grouping within company org charts,
and seniority for tier ordering.

Design principle: rules are checked most-specific → least-specific.
A "Senior Protection & Relay Engineer" should match "Protection & Relay"
before falling through to the generic "Electrical" bucket.
"""

from __future__ import annotations
import json
import re
from pathlib import Path


# ── Helpers ────────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    """Lowercase, collapse whitespace, strip punctuation noise."""
    t = text.lower()
    t = re.sub(r"[,/\-–—|·•]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _any(title: str, keywords: list[str]) -> bool:
    """True if ANY keyword is a substring of the normalized title."""
    t = _norm(title)
    return any(k in t for k in keywords)


def _all(title: str, keywords: list[str]) -> bool:
    """True if ALL keywords appear in the normalized title."""
    t = _norm(title)
    return all(k in t for k in keywords)


# ── DISCIPLINE RULES ───────────────────────────────────────────────────────────
# Ordered most-specific → least-specific within each family.
# Each entry: (discipline_label, match_function)
#
# Discipline labels follow the pattern  "Family — Specialty"
# so the front-end can group by Family or Specialty independently.

DISCIPLINE_RULES: list[tuple[str, callable]] = [

    # ── LEADERSHIP / EXECUTIVE (cross-discipline) ──────────────────────────
    ("Executive",           lambda t: _any(t, ["chief executive", "ceo", "president", "founder", "co-founder"])),
    ("C-Suite — Technical", lambda t: _any(t, ["cto", "coo", "chief technology", "chief operating", "chief engineer"])),
    ("C-Suite — Finance",   lambda t: _any(t, ["cfo", "chief financial"])),
    ("C-Suite — Marketing", lambda t: _any(t, ["cmo", "chief marketing"])),

    # ── PROJECT / PROGRAM MANAGEMENT ──────────────────────────────────────
    ("Project Management — Program Director",    lambda t: _any(t, ["program director", "director of programs"])),
    ("Project Management — Program Manager",     lambda t: _any(t, ["program manager", "pgm"])),
    ("Project Management — Project Manager",     lambda t: _any(t, ["project manager", "pm,", " pm ", "project management"])),
    ("Project Management — Project Engineer",    lambda t: _any(t, ["project engineer"])),
    ("Project Management — Project Controls",    lambda t: _any(t, ["project controls", "cost engineer", "schedule engineer", "planner"])),
    ("Project Management — Construction Mgmt",   lambda t: _any(t, ["construction manager", "construction management", "resident engineer", "field engineer"])),

    # ── COMMISSIONING ─────────────────────────────────────────────────────
    ("Commissioning — Lead Commissioning Eng",   lambda t: _any(t, ["lead commissioning", "commissioning lead", "cx lead"])),
    ("Commissioning — Commissioning Engineer",   lambda t: _any(t, ["commissioning engineer", "cx engineer", "startup engineer", "start-up engineer"])),
    ("Commissioning — Commissioning Technician", lambda t: _any(t, ["commissioning tech", "cx tech", "startup tech"])),
    ("Commissioning — Testing & Inspection",     lambda t: _any(t, ["testing and inspection", "t&i", "test and inspection", "acceptance testing"])),
    ("Commissioning — Functional Testing",       lambda t: _any(t, ["functional test", "fat ", "site acceptance"])),

    # ── ELECTRICAL — fine-grained ─────────────────────────────────────────
    ("Electrical — Protection & Relay",          lambda t: _any(t, ["protection engineer", "relay engineer", "protection & relay", "protection and relay", "relay protection", "p&r ", "protective relay", "distance relay", "overcurrent", "relay coord"])),
    ("Electrical — Substation Design",           lambda t: _any(t, ["substation engineer", "substation design", "substation", "switchyard"]) and not _any(t, ["automation", "scada", "control", "protection"])),
    ("Electrical — Power Systems / Studies",     lambda t: _any(t, ["power systems engineer", "power flow", "load flow", "short circuit", "arc flash", "power quality", "system studies", "grid studies", "power study"])),
    ("Electrical — Transmission",                lambda t: _any(t, ["transmission engineer", "transmission line", "transmission planner", "t-line"])),
    ("Electrical — Distribution",                lambda t: _any(t, ["distribution engineer", "distribution design", "distribution planning", "distribution system"])),
    ("Electrical — Generation",                  lambda t: _any(t, ["generation engineer", "plant electrical", "power plant", "generating station"])),
    ("Electrical — Renewables",                  lambda t: _any(t, ["solar engineer", "wind engineer", "renewable energy engineer", "pv engineer", "bess engineer", "battery storage", "energy storage engineer", "solar", "wind", "renewable"])),
    ("Electrical — High Voltage",                lambda t: _any(t, ["high voltage", "hv engineer", "extra high voltage", "ehv", "345kv", "500kv", "765kv"])),
    ("Electrical — Power Electronics",           lambda t: _any(t, ["power electronics", "inverter", "converter", "vfd", "drives engineer", "ups engineer"])),
    ("Electrical — Instrumentation & Controls",  lambda t: _any(t, ["instrumentation", "i&c", "i & c", "instrument engineer", "control systems engineer", "controls engineer"]) and not _any(t, ["software", "scada", "plc", "dcs"])),
    ("Electrical — SCADA / EMS",                 lambda t: _any(t, ["scada", "ems engineer", "energy management system", "dms engineer", "oms engineer"]) and not _any(t, ["c4isr", "command and control", "surveillance", "reconnaissance"])),
    ("Electrical — Lighting",                    lambda t: _any(t, ["lighting engineer", "lighting designer", "illumination"])),
    ("Electrical — Low Voltage / Building",      lambda t: _any(t, ["low voltage", "building electrical", "mep electrical", "facility electrical"])),
    ("Electrical — General",                     lambda t: _any(t, ["electrical engineer", "electrical designer", "electrician", "electrical technician", "ee,", " ee "])),

    # ── CIVIL — fine-grained ──────────────────────────────────────────────
    ("Civil — Structural",          lambda t: _any(t, ["structural engineer", "structural design", "structures engineer", "se,", " se "])),
    ("Civil — Geotechnical",        lambda t: _any(t, ["geotechnical", "geotech", "soil engineer", "foundation engineer", "geological engineer"])),
    ("Civil — Transportation",      lambda t: _any(t, ["transportation engineer", "traffic engineer", "highway engineer", "roadway engineer", "pavement engineer"])),
    ("Civil — Water Resources",     lambda t: _any(t, ["water resources", "hydraulic engineer", "hydrology", "stormwater", "floodplain", "drainage engineer"])),
    ("Civil — Water / Wastewater",  lambda t: _any(t, ["water engineer", "wastewater engineer", "water treatment", "water infrastructure", "water distribution", "sewer"])),
    ("Civil — Site / Land Dev",     lambda t: _any(t, ["site engineer", "land development", "grading engineer", "site design", "site civil"])),
    ("Civil — Bridge",              lambda t: _any(t, ["bridge engineer", "bridge design", "bridge inspection", "bridge"])),
    ("Civil — General",             lambda t: _any(t, ["civil engineer", "civil designer", "civil technician", "pe,", " pe "])),

    # ── MECHANICAL ────────────────────────────────────────────────────────
    ("Mechanical — HVAC",           lambda t: _any(t, ["hvac", "heating ventilation", "mechanical hvac", "chiller", "cooling engineer"])),
    ("Mechanical — Piping",         lambda t: _any(t, ["piping engineer", "pipe stress", "piping designer", "pipeline engineer", "piping"])),
    ("Mechanical — Rotating Equip", lambda t: _any(t, ["rotating equipment", "turbine engineer", "pump engineer", "compressor engineer", "rotating machinery"])),
    ("Mechanical — Thermodynamics", lambda t: _any(t, ["thermodynamics", "heat transfer", "thermal engineer", "thermal analysis"])),
    ("Mechanical — Structural Mech",lambda t: _any(t, ["fea ", "finite element", "stress analysis", "structural mechanics"])),
    ("Mechanical — Manufacturing",  lambda t: _any(t, ["manufacturing engineer", "process engineer", "tooling engineer", "production engineer", "industrial engineer", "lean engineer", "manufacturing"]) and not _any(t, ["software", "chemical"])),
    ("Mechanical — General",        lambda t: _any(t, ["mechanical engineer", "mechanical designer", "mechanical technician", "me,", " me "])),

    # ── SOFTWARE / CONTROLS ───────────────────────────────────────────────
    ("Software — Embedded Systems",  lambda t: _any(t, ["embedded engineer", "embedded software", "firmware engineer", "real-time software", "rtos", "embedded systems"])),
    ("Software — PLC / DCS",         lambda t: _any(t, ["plc engineer", "dcs engineer", "plc programmer", "dcs programmer", "plc", "dcs", "automation engineer", "automation programmer"])),
    ("Software — Control Systems",   lambda t: _any(t, ["control systems engineer", "controls software", "control software", "feedback control", "pid engineer"])),
    ("Software — Robotics",          lambda t: _any(t, ["robotics engineer", "robotics software", "robot engineer", "ros engineer", "motion planning", "manipulation engineer"])),
    ("Software — AI / ML",           lambda t: _any(t, ["machine learning", "ml engineer", "ai engineer", "deep learning", "data scientist", "nlp engineer", "computer vision engineer"])),
    ("Software — Data Engineering",  lambda t: _any(t, ["data engineer", "data pipeline", "etl engineer", "analytics engineer", "data platform"])),
    ("Software — Full Stack",        lambda t: _any(t, ["full stack", "fullstack", "full-stack"])),
    ("Software — Frontend",          lambda t: _any(t, ["frontend", "front end", "front-end", "ui engineer", "ux engineer", "react engineer", "web developer"])),
    ("Software — Backend",           lambda t: _any(t, ["backend", "back end", "back-end", "api engineer", "server engineer", "platform engineer"])),
    ("Software — DevOps / Cloud",    lambda t: _any(t, ["devops", "site reliability", "sre ", "cloud engineer", "infrastructure engineer", "platform engineer", "devsecops"])),
    ("Software — Cybersecurity",     lambda t: _any(t, ["cybersecurity", "security engineer", "infosec", "penetration", "network security"])),
    ("Software — GIS",               lambda t: _any(t, ["gis engineer", "geospatial engineer", "gis analyst", "gis developer", " gis "])),
    ("Software — General",           lambda t: _any(t, ["software engineer", "software developer", "software architect", "programmer", "developer"])),

    # ── ENVIRONMENTAL / GEOTECHNICAL ──────────────────────────────────────
    ("Environmental — Remediation",  lambda t: _any(t, ["remediation engineer", "environmental remediation", "site remediation", "brownfield"])),
    ("Environmental — Air Quality",  lambda t: _any(t, ["air quality", "emissions engineer", "air pollution", "atmospheric scientist"])),
    ("Environmental — Geotechnical", lambda t: _any(t, ["geotechnical engineer", "geotech engineer", "geotechnical", "soil mechanics"])),
    ("Environmental — Water Quality",lambda t: _any(t, ["water quality", "environmental water"])),
    ("Environmental — Compliance",   lambda t: _any(t, ["environmental compliance", "nepa", "permitting engineer", "environmental permitting", "environmental review"])),
    ("Environmental — General",      lambda t: _any(t, ["environmental engineer", "environmental scientist", "environmental consultant"])),

    # ── CHEMICAL / PROCESS / OIL & GAS ───────────────────────────────────
    ("Chemical — Process Engineering",  lambda t: _any(t, ["process engineer", "process design", "chemical process", "process safety"])),
    ("Chemical — Refining",             lambda t: _any(t, ["refinery engineer", "refining engineer", "petroleum engineer", "downstream engineer"])),
    ("Chemical — Oil & Gas",            lambda t: _any(t, ["oil and gas", "oil & gas", "upstream engineer", "drilling engineer", "reservoir engineer", "completion engineer", "wellbore"])),
    ("Chemical — General",              lambda t: _any(t, ["chemical engineer", "chemist", "biochemical", "materials engineer", "metallurgical engineer", "materials science"])),

    # ── ROBOTICS / AUTOMATION (non-software) ──────────────────────────────
    ("Robotics — Mechanical Design",    lambda t: _any(t, ["robot mechanical", "robotic mechanism", "actuator design", "end effector"])),
    ("Robotics — Integration",          lambda t: _any(t, ["robot integration", "robotics integration", "systems integrator", "robotic cell"])),
    ("Robotics — General",              lambda t: _any(t, ["robotics engineer", "robot engineer", "automation engineer", "robotics"])),

    # ── DEFENSE ───────────────────────────────────────────────────────────
    ("Defense — Systems Engineering",   lambda t: _any(t, ["systems engineer", "systems engineering", "mbse", "model based systems"])),
    ("Defense — Weapons Systems",       lambda t: _any(t, ["weapons system", "weapon system", "ordnance", "munitions engineer", "ballistics"])),
    ("Defense — C4ISR",                 lambda t: _any(t, ["c4isr", "command and control", "intelligence systems", "surveillance engineer", "reconnaissance"])),
    ("Defense — Aerospace / Avionics",  lambda t: _any(t, ["aerospace engineer", "avionics engineer", "flight systems", "aircraft engineer", "aeronautical"])),
    ("Defense — General",               lambda t: _any(t, ["defense engineer", "defence engineer", "DoD", "dod ", "military engineer", "naval engineer", "army engineer", "contractor engineer"])),

    # ── MANUFACTURING (non-mechanical) ────────────────────────────────────
    ("Manufacturing — Quality",         lambda t: _any(t, ["quality engineer", "qa engineer", "qc engineer", "quality assurance", "quality control", "six sigma", "quality manager"])),
    ("Manufacturing — Supply Chain",    lambda t: _any(t, ["supply chain engineer", "procurement engineer", "sourcing engineer", "logistics engineer"])),
    ("Manufacturing — General",         lambda t: _any(t, ["manufacturing engineer", "production engineer", "industrial engineer", "plant engineer", "operations engineer", "manufacturing"])),

    # ── CONSULTING (cross-discipline) ─────────────────────────────────────
    ("Consulting — Partner",            lambda t: _any(t, ["partner,", "partner ", "managing partner", "equity partner", "advisory partner"])),
    ("Consulting — Management",         lambda t: _any(t, ["management consultant", "strategy consultant", "management consulting", "strategy consulting", "business analyst"])),
    ("Consulting — Technical",          lambda t: _any(t, ["technical consultant", "engineering consultant", "solutions consultant", "technical advisor"])),
    ("Consulting — General",            lambda t: _any(t, ["consultant", "consulting"])),

    # ── FINANCE ───────────────────────────────────────────────────────────
    ("Finance — Investment Banking",    lambda t: _any(t, ["investment banker", "investment banking", "ib analyst", "m&a analyst"])),
    ("Finance — Asset Management",      lambda t: _any(t, ["portfolio manager", "asset manager", "fund manager", "investment manager"])),
    ("Finance — General",               lambda t: _any(t, ["financial analyst", "finance analyst", "accountant", "cpa ", "controller", "treasurer", "finance manager"])),

    # ── BUSINESS / OPERATIONS ─────────────────────────────────────────────
    ("Business — Sales & BD",           lambda t: _any(t, ["sales engineer", "business development", "account executive", "account manager", "sales manager", "bd manager"])),
    ("Business — Operations",           lambda t: _any(t, ["operations manager", "operations director", "operations analyst", "chief of staff"])),
    ("Business — HR / People",          lambda t: _any(t, ["human resources", "hr manager", "people operations", "talent acquisition", "recruiter"])),
    ("Business — Marketing",            lambda t: _any(t, ["marketing manager", "marketing analyst", "product marketing", "brand manager", "content strategist"])),
    ("Business — Product Management",   lambda t: _any(t, ["product manager", "product owner", "product lead", " pm,", "head of product"])),
    ("Business — General",              lambda t: _any(t, ["director", "vice president", "vp ", "svp", "evp", "managing director"])),

    # ── ACADEMIA / RESEARCH ───────────────────────────────────────────────
    ("Academia — Research",             lambda t: _any(t, ["research engineer", "research scientist", "researcher", "r&d engineer", "r & d"])),
    ("Academia — Professor",            lambda t: _any(t, ["professor", "lecturer", "faculty", "academic", "postdoc", "phd researcher"])),

    # ── CATCH-ALL ─────────────────────────────────────────────────────────
    ("Unknown", lambda t: True),
]


# ── SENIORITY RULES ────────────────────────────────────────────────────────────
# Checked independently of discipline.
# Ordered most-specific → least-specific.
#
# PE (Professional Engineer) is a CERTIFICATION, not a seniority level.
# It is detected separately via the `is_pe` field in classify() and exposed
# as a boolean for UI filtering — it does not affect seniority tier placement.
#
# EIT (Engineer in Training) IS its own tier — between Entry-level and Intern.
#
# Grade numbers (Engineer I/II/III/IV/V) get their own explicit tiers so
# an "Engineer III" is never lumped with a "Senior Engineer".

SENIORITY_RULES: list[tuple[str, callable]] = [

    # ── Must come FIRST — intern/EIT before anything else ─────────────────
    # "intern" contains "in" which would otherwise match "principal "
    ("Intern / Co-op",  lambda t: _any(t, [
        "intern", "internship", "co-op", "coop", "co op",
        "student engineer", "student worker", "student employee",
        "graduate trainee", "graduate student",
    ])),
    ("EIT",             lambda t: _any(t, [
        "engineer in training", "eit", "engineer-in-training",
        "engineering intern", "engineering trainee",
        "graduate engineer",   # common UK/AU title for new grads
    ])),

    # ── Executive ──────────────────────────────────────────────────────────
    ("Partner / Principal (Consulting)", lambda t: _any(t, [
        "managing partner", "equity partner", "principal consultant",
    ]) or (_any(t, ["partner"]) and not _any(t, ["intern", "eit"]))),
    ("Fellow / Distinguished",  lambda t: _any(t, [
        "fellow", "distinguished engineer", "distinguished scientist", "ieee fellow",
    ])),
    ("C-Suite",                 lambda t: _any(t, [
        "chief executive", "chief operating", "chief financial",
        "chief technology", "chief marketing", "chief engineer",
        "ceo", "coo", "cfo", "cto", "cmo", "president", "founder", "co-founder",
    ])),

    # ── VP ─────────────────────────────────────────────────────────────────
    ("EVP / SVP",   lambda t: _any(t, ["executive vice president", "evp", "senior vice president", "svp"])),
    ("VP",          lambda t: _any(t, ["vice president", " vp ", "vp,", "vp-", "vp of"])),
    ("AVP",         lambda t: _any(t, ["assistant vice president", "avp"])),

    # ── Director ───────────────────────────────────────────────────────────
    ("Director",    lambda t: _any(t, ["director", "managing director"])),

    # ── Management ─────────────────────────────────────────────────────────
    ("Group Manager / Dept Head", lambda t: _any(t, [
        "group manager", "department head", "department manager",
        "division manager", "practice leader", "practice manager",
    ])),
    ("Manager",     lambda t: _any(t, ["manager", "head of ", "team lead,"])),

    # ── Technical leadership ───────────────────────────────────────────────
    ("Principal Engineer",  lambda t: _any(t, [
        "principal engineer", "staff engineer", "distinguished engineer",
        "principal ", "staff ",  # catches "Principal Structural Engineer" etc.
    ]) and not _any(t, ["intern", "eit", "principal consultant", "managing partner"])),
    ("Lead Engineer",       lambda t: _any(t, [
        "lead engineer", "engineering lead", "technical lead", "tech lead",
    ])),

    # ── Numbered grade tiers (most specific — checked before Senior/Mid) ───
    # Uses regex to catch both "Engineer III" at end-of-string and mid-title.
    ("Engineer V",   lambda t: bool(re.search(r"(engineer|eng|analyst|scientist|designer).{0,6}\b(v|5)\b", _norm(t))) or
                               bool(re.search(r"\b(level|grade|e)[ -]?(v|5)\b", _norm(t)))),
    ("Engineer IV",  lambda t: bool(re.search(r"(engineer|eng|analyst|scientist|designer).{0,6}\b(iv|4)\b", _norm(t))) or
                               bool(re.search(r"\b(level|grade|e)[ -]?(iv|4)\b", _norm(t)))),
    ("Engineer III", lambda t: bool(re.search(r"(engineer|eng|analyst|scientist|designer).{0,6}\b(iii|3)\b", _norm(t))) or
                               bool(re.search(r"\b(level|grade|e)[ -]?(iii|3)\b", _norm(t)))),
    ("Engineer II",  lambda t: bool(re.search(r"(engineer|eng|analyst|scientist|designer).{0,6}\b(ii|2)\b", _norm(t))) or
                               bool(re.search(r"\b(level|grade|e)[ -]?(ii|2)\b", _norm(t)))),
    ("Engineer I",   lambda t: bool(re.search(r"(engineer|eng|analyst|scientist|designer).{0,6}\b(i|1)\b", _norm(t))) or
                               bool(re.search(r"\b(level|grade|e)[ -]?(i|1)\b", _norm(t)))),

    # ── Word-based seniority (after numbered grades) ───────────────────────
    ("Senior",      lambda t: _any(t, [
        "senior", "sr.", " sr ", "sr,",
    ]) and not _any(t, ["intern", "eit"])),

    ("Mid-level",   lambda t: _any(t, [
        "associate engineer", "associate analyst",
    ])),

    ("Entry-level", lambda t: _any(t, [
        "junior", "jr.", " jr ", "jr,", "entry level", "entry-level",
    ]) and not _any(t, ["intern", "eit"])),

    # ── Project / Commissioning ────────────────────────────────────────────
    ("Project Engineer",  lambda t: _any(t, ["project engineer"])),
    ("Commissioning Eng", lambda t: _any(t, [
        "commissioning engineer", "cx engineer", "startup engineer", "start-up engineer",
    ])),

    # ── Technicians / drafters ─────────────────────────────────────────────
    ("Technician",  lambda t: _any(t, [
        "technician", "drafter", "cad ", "cadd ",
        "designer", "detailer",
    ])),

    # ── Catch-all — matched a discipline but no specific seniority ─────────
    ("Engineer",    lambda t: _any(t, [
        "engineer", "scientist", "analyst", "developer",
        "consultant", "specialist", "architect",
    ])),

    ("Unknown",     lambda t: True),
]

# ── PE certification detector ──────────────────────────────────────────────────
# Returns True if the title suggests a PE license.
# Used by classify() to set the `is_pe` boolean field.
_PE_PATTERNS = re.compile(
    r"(pe|p\.e\.|professional engineer|licensed engineer|registered engineer)",
    re.IGNORECASE,
)

def detect_pe(raw_title: str) -> bool:
    """Return True if the title contains a PE certification indicator."""
    if not raw_title:
        return False
    return bool(_PE_PATTERNS.search(raw_title))


# ── Custom taxonomy loader ────────────────────────────────────────────────────

def _load_custom_taxonomy() -> dict:
    """
    Load custom_taxonomy.json from the data root defined in config.py.
    Returns empty structure on any error so the built-in rules still work.
    """
    try:
        # Import here to avoid circular imports at module load time
        from config import CUSTOM_TAXONOMY_PATH
        custom_path = CUSTOM_TAXONOMY_PATH
    except Exception:
        # Fallback if config isn't available (e.g. running taxonomy standalone)
        custom_path = Path(__file__).parent / "data" / "custom_taxonomy.json"

    try:
        if custom_path.exists():
            with open(custom_path, encoding="utf-8") as f:
                data = json.load(f)
            return data
    except Exception as e:
        print(f"[taxonomy] Warning: could not load custom_taxonomy.json: {e}")
    return {}

def _build_custom_rule(keywords: list[str]):
    """Turn a list of keyword strings into a lambda matching any of them."""
    kws = [k.lower() for k in keywords if isinstance(k, str) and not k.startswith("_")]
    return lambda t: any(kw in _norm(t) for kw in kws)

def _get_custom_rules() -> tuple[list, list, list]:
    """
    Returns (custom_discipline_rules, custom_seniority_rules, certifications).
    custom_discipline_rules: [(label, fn), ...] — prepended to DISCIPLINE_RULES
    custom_seniority_rules:  [(label, fn), ...] — prepended to SENIORITY_RULES
    certifications:          [{name, fn}, ...]  — checked in classify()
    """
    data = _load_custom_taxonomy()

    disc_rules = []
    for entry in data.get("discipline_rules", []):
        label = entry.get("label", "")
        kws   = entry.get("keywords", [])
        if label and kws:
            disc_rules.append((label, _build_custom_rule(kws)))

    seniority_rules = []
    for entry in data.get("seniority_rules", []):
        label = entry.get("label", "")
        kws   = entry.get("keywords", [])
        if label and kws:
            seniority_rules.append((label, _build_custom_rule(kws)))

    certs = []
    for entry in data.get("certifications", []):
        name = entry.get("name", "")
        kws  = entry.get("keywords", [])
        if name and kws:
            certs.append({"name": name, "fn": _build_custom_rule(kws)})

    return disc_rules, seniority_rules, certs


# ── PUBLIC API ─────────────────────────────────────────────────────────────────

def classify(raw_title: str) -> dict[str, str]:
    """
    Classify a raw LinkedIn job title.

    Returns:
      {
        "discipline": "Electrical — Protection & Relay",
        "discipline_family": "Electrical",
        "discipline_specialty": "Protection & Relay",
        "seniority": "Senior",
        "label": "Senior Electrical — Protection & Relay Engineer"
      }

    Always returns a result — falls back to "Unknown" fields rather than raising.
    """
    if not raw_title or not raw_title.strip():
        return _make_result("Unknown", "Unknown")

    # Load custom rules fresh each call so edits to custom_taxonomy.json
    # are picked up without restarting the program.
    # (Cached after first load for performance — restart to reload.)
    custom_disc, custom_seniority, custom_certs = _get_custom_rules()

    # Custom discipline rules are checked FIRST so they can override built-ins
    discipline = "Unknown"
    for label, fn in custom_disc + DISCIPLINE_RULES:
        try:
            if fn(raw_title):
                discipline = label
                break
        except Exception:
            continue

    # Custom seniority rules are checked FIRST
    seniority = "Unknown"
    for label, fn in custom_seniority + SENIORITY_RULES:
        try:
            if fn(raw_title):
                seniority = label
                break
        except Exception:
            continue

    # Built-in PE detection
    is_pe = detect_pe(raw_title)

    # Custom certifications
    certifications = {}
    for cert in custom_certs:
        try:
            certifications[cert["name"]] = cert["fn"](raw_title)
        except Exception:
            certifications[cert["name"]] = False

    return _make_result(discipline, seniority, is_pe, certifications)


def _make_result(discipline: str, seniority: str, is_pe: bool = False, certifications: dict = None) -> dict:
    # Split "Family — Specialty" into parts
    if " — " in discipline:
        family, specialty = discipline.split(" — ", 1)
    else:
        family, specialty = discipline, ""

    label = f"{seniority} · {discipline}" if specialty else f"{seniority} · {family}"

    return {
        "discipline":           discipline,
        "discipline_family":    family,
        "discipline_specialty": specialty,
        "seniority":            seniority,
        "is_pe":                is_pe,
        "certifications":       certifications or {},
        "label":                label,
    }


# ── CLI quick-test ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    test_titles = [
        "Senior Protection & Relay Engineer",
        "Staff Software Engineer, Embedded Systems",
        "VP of Infrastructure",
        "Civil Engineer III – Geotechnical",
        "Civil Engineer II",
        "Civil Engineer I",
        "Engineer III",
        "Lead Commissioning Engineer",
        "Robotics Engineer – Motion Planning",
        "Senior PLC/DCS Automation Engineer",
        "Project Manager – Power Delivery",
        "Principal Structural Engineer",
        "Machine Learning Engineer",
        "Defense Systems Engineer (C4ISR)",
        "Manufacturing Quality Engineer – Six Sigma",
        "Environmental Compliance Specialist, NEPA",
        "Partner, Infrastructure Advisory",
        "Electrical Designer – Substation",
        "SCADA Engineer",
        "Geotechnical Engineer II",
        "Engineering Intern",
        "Intern",
        "Summer Intern",
        "Co-op Student",
        "Engineer in Training",
        "EIT",
        "Graduate Engineer",
        "Junior Civil Engineer",
        "PE, Senior Structural Engineer",
        "Licensed Professional Engineer",
        "",
    ]
    print(f"{'Raw Title':<45}  {'Discipline':<30}  {'Seniority':<22}  PE?")
    print("─" * 110)
    for t in test_titles:
        r = classify(t)
        pe = "✓ PE" if r["is_pe"] else ""
        print(f"{(t or '(empty)'):<45}  {r['discipline']:<30}  {r['seniority']:<22}  {pe}")
