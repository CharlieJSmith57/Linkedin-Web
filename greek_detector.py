"""
greek_detector.py  —  Greek org & chapter detection for the LinkedIn enricher

HOW IT WORKS
────────────
1. At import time, load_greek_index() builds a lookup dictionary from all 7
   batch JSON files.  The index maps every known alias (org full name, letters,
   common abbreviations, and chapter+school combinations) → org metadata.

2. detect_greek_orgs(profile_data) scans every text field in a LinkedIn profile
   using a tiered approach:
     a. Org full-name substrings  (e.g., "Theta Chi", "Sigma Alpha Epsilon")
     b. Greek letter sequences    (e.g., "SAE", "TKE", "Alpha Phi Alpha")
     c. Chapter+org co-occurrence (e.g., "Epsilon Chapter of Theta Chi")
     d. Unicode Greek letters     (e.g., "ΘΧ", "ΣΑΕ")

3. Returns a deduplicated list of match dicts (one per org_id).

USAGE IN ingest_enriched.py
───────────────────────────
    from greek_detector import detect_greek_orgs
    matches = detect_greek_orgs(raw_profile)
    # store as JSON string in the connection record

ADDING NEW ORGS
───────────────
Drop a new greek_orgs_batch_N.json into B:\\linkedin-data\\data\\ — the
detector auto-discovers all files matching greek_orgs_batch_*.json.
"""

from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# ── Configuration ─────────────────────────────────────────────────────────────

_DEFAULT_DATA_DIR = Path(os.environ.get(
    "LINKEDIN_DATA_DIR",
    r"B:\linkedin-data\data"
))

# ── Index types ────────────────────────────────────────────────────────────────

_OrgMeta     = dict   # org_id, full_name, letters, active_chapters
_ChapterMeta = dict   # org_id, full_name, letters, chapter, school, abbrev, status


class GreekIndex:
    """Compiled lookup structure from all batch JSON files."""

    def __init__(self) -> None:
        self._org_by_id: dict[str, _OrgMeta] = {}
        # normalized alias → list of OrgMeta
        self._org_name_index: dict[str, list[_OrgMeta]] = {}
        # normalized letters → list of OrgMeta
        self._letters_index: dict[str, list[_OrgMeta]] = {}
        # (norm_org_name, norm_chapter) → list of ChapterMeta
        self._org_chapter_index: dict[tuple, list[_ChapterMeta]] = {}
        # org_id → list of all ChapterMeta (for reverse lookup)
        self._chapters_by_org: dict[str, list[_ChapterMeta]] = {}
        self._loaded_batches: list[str] = []

    # ── Loading ────────────────────────────────────────────────────────────────

    def load_all(self, data_dir: Path) -> None:
        batch_files = sorted(data_dir.glob("greek_orgs_batch_*.json"))
        if not batch_files:
            logger.warning("greek_detector: no batch files in %s", data_dir)
            return
        for path in batch_files:
            self._load_batch(path)
        logger.info(
            "greek_detector: %d orgs, %d org-chapter pairs, %d batch files",
            len(self._org_by_id),
            sum(len(v) for v in self._org_chapter_index.values()),
            len(self._loaded_batches),
        )

    def _load_batch(self, path: Path) -> None:
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as exc:
            logger.error("greek_detector: failed to load %s — %s", path, exc)
            return
        self._loaded_batches.append(path.name)

        for org in data.get("orgs", []):
            org_id   = org["id"]
            full     = org["full_name"]
            letters  = org.get("letters", "")
            active   = org.get("active_chapters", 0)

            meta: _OrgMeta = {
                "org_id": org_id, "full_name": full,
                "letters": letters, "active_chapters": active,
            }
            self._org_by_id[org_id] = meta

            # Index by full name and all aliases
            # Skip single-Greek-letter-word keys (too ambiguous — "beta" matches everywhere)
            _SINGLE_GREEK_NORM = {
                'alpha','beta','gamma','delta','epsilon','zeta','eta','theta',
                'iota','kappa','lambda','mu','nu','xi','omicron','pi','rho',
                'sigma','tau','upsilon','phi','chi','psi','omega',
            }
            for alias in [full] + _id_to_aliases(org_id, full):
                key = _norm(alias)
                if key and len(key) >= 4 and key not in _SINGLE_GREEK_NORM:
                    self._org_name_index.setdefault(key, []).append(meta)

            # Index by letters
            if letters:
                key = _norm(letters)
                if key:
                    self._letters_index.setdefault(key, []).append(meta)

            # Build chapter index keyed by (norm_org_name_alias, norm_chapter_name)
            # This ensures chapter only fires when the ORG name is also present
            all_org_keys: list[str] = []
            for alias in [full] + _id_to_aliases(org_id, full):
                k = _norm(alias)
                if k and len(k) >= 4:
                    all_org_keys.append(k)
            if letters:
                k = _norm(letters)
                if k:
                    all_org_keys.append(k)

            for chap in org.get("chapters", []):
                chap_name = chap["chapter"]
                school    = chap.get("school", "")
                abbrev    = chap.get("abbrev", "")
                status    = chap.get("status", "active")

                cm: _ChapterMeta = {
                    "org_id": org_id, "full_name": full, "letters": letters,
                    "chapter": chap_name, "school": school,
                    "abbrev": abbrev, "status": status,
                }

                self._chapters_by_org.setdefault(org_id, []).append(cm)

                nchap = _norm(chap_name)
                for org_key in all_org_keys:
                    pair = (org_key, nchap)
                    self._org_chapter_index.setdefault(pair, []).append(cm)

    # ── Lookups ────────────────────────────────────────────────────────────────

    def lookup_org_name(self, key: str) -> list[_OrgMeta]:
        return self._org_name_index.get(key, [])

    def lookup_letters(self, key: str) -> list[_OrgMeta]:
        return self._letters_index.get(key, [])

    def lookup_org_chapter(self, org_key: str, chapter_key: str) -> list[_ChapterMeta]:
        """Only matches when the specific org is also identified in the text."""
        return self._org_chapter_index.get((org_key, chapter_key), [])

    def all_org_keys(self) -> list[str]:
        return list(self._org_name_index.keys())

    def all_letter_keys(self) -> list[str]:
        return list(self._letters_index.keys())


# ── Singleton ──────────────────────────────────────────────────────────────────

_INDEX: GreekIndex | None = None


def load_greek_index(data_dir: Path | None = None) -> GreekIndex:
    global _INDEX
    if _INDEX is None:
        _INDEX = GreekIndex()
        _INDEX.load_all(data_dir or _DEFAULT_DATA_DIR)
    return _INDEX


def reload_greek_index(data_dir: Path | None = None) -> GreekIndex:
    global _INDEX
    _INDEX = None
    return load_greek_index(data_dir)


# ── Main detection function ────────────────────────────────────────────────────

def detect_greek_orgs(
    profile: dict[str, Any],
    data_dir: Path | None = None,
) -> list[dict[str, Any]]:
    """
    Scan a raw LinkedIn profile dict for Greek org affiliations.
    Returns deduplicated list of match dicts (one per org_id).
    """
    index = load_greek_index(data_dir)
    matches: dict[str, dict] = {}

    texts = _extract_text_fields(profile)

    for raw_text, field_name in texts:
        if not raw_text:
            continue
        _scan_text(raw_text, field_name, index, matches)

    return sorted(matches.values(), key=lambda m: m["full_name"])


# ── Text extraction ────────────────────────────────────────────────────────────

def _extract_text_fields(profile: dict) -> list[tuple[str, str]]:
    texts: list[tuple[str, str]] = []

    def add(v: Any, label: str) -> None:
        if isinstance(v, str) and v.strip():
            texts.append((v.strip(), label))

    add(profile.get("summary"),    "summary")
    add(profile.get("headline"),   "headline")
    add(profile.get("occupation"), "occupation")

    for pos in profile.get("experience", []):
        add(pos.get("companyName"), "pos.company")
        add(pos.get("title"),       "pos.title")
        add(pos.get("description"), "pos.desc")

    for edu in profile.get("education", []):
        add(edu.get("schoolName"),   "edu.school")
        add(edu.get("fieldOfStudy"), "edu.field")
        add(edu.get("description"),  "edu.desc")
        add(edu.get("activities"),   "edu.activities")
        add(edu.get("degreeName"),   "edu.degree")

    for vol in profile.get("volunteer", []):
        add(vol.get("companyName"), "vol.company")
        add(vol.get("role"),        "vol.role")
        add(vol.get("description"), "vol.desc")

    for cert in profile.get("certifications", []):
        add(cert.get("name"),      "cert.name")
        add(cert.get("authority"), "cert.authority")

    for honor in profile.get("honors", []):
        add(honor.get("title"),       "honor.title")
        add(honor.get("description"), "honor.desc")

    for proj in profile.get("projects", []):
        add(proj.get("title"),       "proj.title")
        add(proj.get("description"), "proj.desc")

    for course in profile.get("courses", []):
        add(course.get("name"), "course.name")

    return texts


# ── Scanner ────────────────────────────────────────────────────────────────────

_GREEK_LETTER_WORDS = [
    "alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta",
    "iota", "kappa", "lambda", "mu", "nu", "xi", "omicron", "pi", "rho",
    "sigma", "tau", "upsilon", "phi", "chi", "psi", "omega",
]
_GREEK_WORD_RE = re.compile(
    r"\b(" + "|".join(_GREEK_LETTER_WORDS) + r")\b", re.IGNORECASE
)
_UNICODE_GREEK_RE = re.compile(
    r"[ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"
    r"αβγδεζηθικλμνξοπρστυφχψω]{2,6}"
)


def _scan_text(
    text: str,
    field: str,
    index: GreekIndex,
    matches: dict[str, dict],
) -> None:
    norm = _norm(text)

    # ── Tier 1: org full-name / alias substrings ───────────────────────────────
    # We also collect which org_keys were found (for tier 3 chapter lookup)
    found_org_keys: set[str] = set()

    for key in index.all_org_keys():
        if key not in norm:
            continue
        # Check word-boundary to avoid "alpha" matching "alphanumeric"
        # For multi-word keys (most org names), substring is fine
        if " " not in key:
            # Single-word key — require word boundary
            if not re.search(r"\b" + re.escape(key) + r"\b", norm):
                continue
        for org_meta in index.lookup_org_name(key):
            found_org_keys.add(key)
            _add_match(
                matches, org_meta,
                chapter=None, school=None,
                match_type="org_name",
                evidence=_snippet(text, key),
            )

    # ── Dedup: remove shorter-name matches that are subsumed by longer ones ─────
    # e.g., if "Alpha Phi Alpha" matched, suppress "Alpha Phi" (which is a prefix)
    if len(matches) > 1:
        names_found = {m["full_name"].lower() for m in matches.values()}
        subsumed = set()
        for org_id_a, m_a in matches.items():
            name_a = m_a["full_name"].lower()
            for name_b in names_found:
                if name_b != name_a and name_b.startswith(name_a + " "):
                    subsumed.add(org_id_a)
                    break
        for org_id_rm in subsumed:
            del matches[org_id_rm]

    # ── Tier 2: Greek letter word sequences (2–4 words) ───────────────────────
    words = _GREEK_WORD_RE.findall(text)
    for window in (4, 3, 2):
        for i in range(len(words) - window + 1):
            candidate = " ".join(words[i:i + window])
            nc = _norm(candidate)
            hits = index.lookup_org_name(nc) or index.lookup_letters(nc)
            for org_meta in hits:
                found_org_keys.add(nc)
                _add_match(
                    matches, org_meta,
                    chapter=None, school=None,
                    match_type="letters",
                    evidence=candidate,
                )

    # ── Tier 3: Chapter lookup — only when org already found in this text ──────
    # Pattern: 1–3 Greek-letter words followed by "chapter" or "fraternity/sorority"
    # but ONLY fire if the parent org was already identified in the same text.
    if found_org_keys:
        chapter_re = re.compile(
            r"\b((?:" + "|".join(_GREEK_LETTER_WORDS) + r")(?:\s+(?:"
            + "|".join(_GREEK_LETTER_WORDS) + r")){0,2})"
            r"\s+(?:chapter|fraternity|sorority)\b",
            re.IGNORECASE,
        )
        for m in chapter_re.finditer(text):
            chapter_str = m.group(1).strip()
            nchap = _norm(chapter_str)
            for org_key in found_org_keys:
                for cm in index.lookup_org_chapter(org_key, nchap):
                    org_meta = {
                        "org_id": cm["org_id"], "full_name": cm["full_name"],
                        "letters": cm["letters"], "active_chapters": 0,
                    }
                    _add_match(
                        matches, org_meta,
                        chapter=cm["chapter"], school=cm["school"],
                        match_type="chapter_match",
                        evidence=_snippet(text, chapter_str),
                    )

    # ── Tier 4: Unicode Greek letter strings ──────────────────────────────────
    for m in _UNICODE_GREEK_RE.finditer(text):
        candidate = m.group(0)
        nc = _norm(candidate)
        for org_meta in index.lookup_letters(nc):
            _add_match(
                matches, org_meta,
                chapter=None, school=None,
                match_type="unicode_letters",
                evidence=candidate,
            )


# ── Match management ───────────────────────────────────────────────────────────

_PRIORITY = {
    "chapter_match":   4,
    "org_name":        3,
    "letters":         2,
    "unicode_letters": 2,
    "alias":           1,
}


def _add_match(
    matches: dict,
    org_meta: _OrgMeta,
    chapter: str | None,
    school: str | None,
    match_type: str,
    evidence: str,
) -> None:
    org_id   = org_meta["org_id"]
    priority = _PRIORITY.get(match_type, 0)
    if org_id in matches:
        existing_priority = _PRIORITY.get(matches[org_id].get("match_type", ""), 0)
        if priority <= existing_priority:
            return
    matches[org_id] = {
        "org_id":     org_id,
        "full_name":  org_meta["full_name"],
        "letters":    org_meta["letters"],
        "chapter":    chapter,
        "school":     school,
        "match_type": match_type,
        "evidence":   evidence[:200],
    }


# ── Utilities ─────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    if not text:
        return ""
    t = text.lower()
    t = re.sub(r"[,\.\-–—|·•/\\'\"]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _snippet(text: str, key: str, context: int = 80) -> str:
    idx = _norm(text).find(_norm(key))
    if idx < 0:
        return text[:context]
    start = max(0, idx - context // 2)
    end   = min(len(text), idx + len(key) + context // 2)
    return text[start:end].strip()


def _id_to_aliases(org_id: str, full_name: str) -> list[str]:
    aliases: list[str] = []
    from_id = org_id.replace("_", " ").title()
    if from_id.lower() != full_name.lower():
        aliases.append(from_id)
    _SINGLE_GREEK = {
        'alpha','beta','gamma','delta','epsilon','zeta','eta','theta',
        'iota','kappa','lambda','mu','nu','xi','omicron','pi','rho',
        'sigma','tau','upsilon','phi','chi','psi','omega',
    }
    if len(org_id) <= 5 and org_id not in _SINGLE_GREEK:
        aliases.append(org_id.upper())

    _COLLOQUIALS: dict[str, list[str]] = {
        "fiji":              ["FIJI", "Phi Gam"],
        "pike":              ["PIKE"],
        "tke":               ["TKE"],
        "sig_ep":            ["SigEp", "Sig Ep"],
        "sig_chi":           ["Sig Chi"],
        "lambda_chi":        ["Lambda Chi", "LCA"],
        "ato":               ["ATO"],
        "phi_delt":          ["Phi Delt"],
        "aepi":              ["AEPi"],
        "delt":              ["Delt", "DTD"],
        "sigma_nu":          ["Sig Nu"],
        "pi_kapp":           ["Pi Kapp", "PKP"],
        "phi_psi":           ["Phi Psi"],
        "phi_kap_tau":       ["PhiTau", "Phi Tau"],
        "dke":               ["DKE", "Dekes"],
        "alpha_sig":         ["Alpha Sig"],
        "kkg":               ["KKG"],
        "tri_delt":          ["TriDelt", "Tri Delt", "DDD"],
        "adpi":              ["ADPi", "AD Pi"],
        "alpha_phi":         ["APhi", "A Phi"],
        "pi_phi":            ["PiPhi", "Pi Phi"],
        "dg":                ["DG"],
        "kd":                ["KD"],
        "zta":               ["ZTA"],
        "dz":                ["DZ"],
        "gamma_phi":         ["GPhi", "G Phi B"],
        "kappa_alpha_theta": ["KAT"],
        "alpha_kappa_alpha": ["AKA"],
        "kappa_alpha_psi":   ["Kappas", "Nupes"],
        "delta_sigma_theta": ["Deltas", "DST"],
        "phi_beta_sigma":    ["Sigmas", "PBS"],
        "omega_psi_phi":     ["Ques", "OPP"],
        "zeta_phi_beta":     ["Zetas", "ZPhiB"],
        "sigma_gamma_rho":   ["SGRho"],
        "tau_beta_pi":       ["TBP"],
        "phi_sigma_rho":     ["Phi Sig Rho"],
        "kappa_alpha_order": ["Kappa Alpha", "KA Order"],
        "theta_tau":         ["Theta Tau Fraternity"],
        "sigma_phi_delta":   ["Sig Phi Delta"],
        "farm_house":        ["FarmHouse Fraternity"],
        "acacia":            ["Acacia Fraternity"],
        "delta_upsilon":     ["DU", "Delta U"],
        "sigma_kappa":       ["SK", "Sig Kap"],
        "alpha_chi_omega":   ["AXO", "Alpha Chi"],
        "alpha_omicron_pi":  ["AOII", "AOPi", "Alpha O"],
        "alpha_gamma_delta": ["AGD", "Alpha Gam"],
        "phi_mu":            ["Phi Mu Fraternity"],
        "phi_sigma_sigma":   ["Phi Sig Sigma", "PhiSigSig"],
        "alpha_xi_delta":    ["AXiD", "Alpha Xi"],
        "alpha_epsilon_phi": ["AEPhi"],
        "sigma_delta_tau":   ["SDT", "Sig Delt Tau"],
        "delta_phi_epsilon": ["DPhiE", "D Phi E"],
        "iota_phi_theta":    ["Iotas"],
        "theta_delta_chi":   ["TDC", "TDX"],
        "kappa_alpha_order": ["Kappa Alpha Order"],
    }
    for extra in _COLLOQUIALS.get(org_id, []):
        aliases.append(extra)
    return aliases


# ── Convenience: tag a connection record ──────────────────────────────────────

def tag_connection(
    record: dict[str, Any],
    data_dir: Path | None = None,
) -> dict[str, Any]:
    """
    Detect Greek orgs in a connection record and write results back.
    Adds: greek_orgs (JSON), greek_org_ids (CSV), greek_org_names (CSV).
    """
    raw_profile = record.get("raw_profile") or _flatten_to_profile(record)
    matches = detect_greek_orgs(raw_profile, data_dir)
    record["greek_orgs"]      = json.dumps(matches, ensure_ascii=False)
    record["greek_org_ids"]   = ", ".join(m["org_id"]    for m in matches)
    record["greek_org_names"] = ", ".join(m["full_name"] for m in matches)
    return record


def _flatten_to_profile(record: dict) -> dict:
    positions = []
    if record.get("current_company"):
        positions.append({
            "companyName": record.get("current_company", ""),
            "title":       record.get("current_title", ""),
            "description": "",
        })
    for i in range(1, 11):
        co = record.get(f"prev_company_{i}")
        if co:
            positions.append({
                "companyName": co,
                "title":       record.get(f"prev_title_{i}", ""),
                "description": "",
            })
    return {
        "headline":   record.get("headline", ""),
        "summary":    "",
        "experience": positions,
        "education":  [{"schoolName": record.get("school", "")}]
                      if record.get("school") else [],
        "volunteer":  [],
    }


# ── CLI quick-test ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO)
    data_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else None

    idx = load_greek_index(data_dir)
    print(f"Index: {len(idx._org_by_id)} orgs, "
          f"{len(idx._org_name_index)} name keys, "
          f"{len(idx._org_chapter_index)} org+chapter pairs\n")

    test_profiles = [
        {
            "description": "Theta Chi – Epsilon Chapter (WPI)",
            "profile": {
                "headline": "Electrical Engineer at Jacobs",
                "experience": [
                    {"companyName": "Theta Chi Fraternity", "title": "President",
                     "description": "Served as president of Epsilon Chapter at WPI"},
                    {"companyName": "Jacobs Engineering", "title": "EE", "description": ""},
                ],
                "education": [{"schoolName": "Worcester Polytechnic Institute",
                               "activities": "Theta Chi, IEEE, FIRST Robotics"}],
                "volunteer": [], "summary": "",
            }
        },
        {
            "description": "SAE alumnus, PE-licensed engineer",
            "profile": {
                "headline": "PE | Senior Protection & Relay Engineer | SAE Alumni",
                "experience": [
                    {"companyName": "Sigma Alpha Epsilon – Massachusetts Delta",
                     "title": "Risk Manager", "description": "SAE chapter operations"},
                ],
                "education": [{"schoolName": "WPI", "activities": "SAE, SWE"}],
                "volunteer": [], "summary": "",
            }
        },
        {
            "description": "NPHC — Alpha Phi Alpha + mention of Kappa Alpha Psi",
            "profile": {
                "headline": "Software Engineer | Alpha Phi Alpha",
                "experience": [{"companyName": "Google", "title": "SWE", "description": ""}],
                "education": [{"schoolName": "Howard University",
                               "activities": "ΑΦΑ, NSBE, Interest in Kappa Alpha Psi"}],
                "volunteer": [], "summary": "",
            }
        },
        {
            "description": "Tau Beta Pi + Triangle (engineering fraternitites)",
            "profile": {
                "headline": "Civil Engineer | Tau Beta Pi | Triangle Fraternity",
                "experience": [
                    {"companyName": "Triangle Fraternity – Missouri S&T",
                     "title": "VP", "description": "Engineering fraternity leadership"},
                ],
                "education": [], "volunteer": [], "summary": "",
            }
        },
        {
            "description": "No greek org (should return empty)",
            "profile": {
                "headline": "Project Manager at AECOM",
                "experience": [{"companyName": "AECOM", "title": "PM", "description": ""}],
                "education": [{"schoolName": "Georgia Tech", "activities": "SWE, ASCE"}],
                "volunteer": [], "summary": "",
            }
        },
        {
            "description": "Phi Sigma Rho (engineering sorority)",
            "profile": {
                "headline": "EE at Eaton | Phi Sigma Rho",
                "experience": [
                    {"companyName": "Phi Sigma Rho Sorority", "title": "President",
                     "description": "Women in engineering sorority"},
                ],
                "education": [{"schoolName": "Purdue University",
                               "activities": "Phi Sigma Rho, SWE, IEEE"}],
                "volunteer": [], "summary": "",
            }
        },
    ]

    for t in test_profiles:
        results = detect_greek_orgs(t["profile"], data_dir)
        print(f"── {t['description']}")
        if results:
            for r in results:
                chap = f"  ({r['chapter']} @ {r['school']})" if r.get("chapter") else ""
                print(f"  ✓ {r['full_name']} [{r['letters']}]{chap}  [{r['match_type']}]")
        else:
            print("  (no matches)")
        print()
