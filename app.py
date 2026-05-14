"""
HVPQ / PIQ Vetting Observation Checker
--------------------------------------
Rule-based Streamlit app for objective, exportable HVPQ/PIQ checks.

Design principles:
1. Simple and objective checks first.
2. Do not infer from JiBe incident list. Only compare HVPQ/PIQ declarations and flag blanks/no incidents.
3. Class Status formats differ by class society, so use generic proximity-based extraction and mark uncertain items for manual check.
4. Output is a targeted discrepancy/manual verification register that can be sent to ship.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, asdict
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st
from dateutil import parser as dateparser
from dateutil.relativedelta import relativedelta
from rapidfuzz import fuzz

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

try:
    from docx import Document
except Exception:  # pragma: no cover
    Document = None


# -----------------------------
# Configuration
# -----------------------------

APP_TITLE = "HVPQ / PIQ Vetting Observation Checker"
DEFAULT_CERT_WARNING_DAYS = 90
DEFAULT_SURVEY_WARNING_DAYS = 180
DEFAULT_TS_GAP_MONTHS = 7.0
DEFAULT_MS_GAP_MONTHS = 12.0

RISK_SCORE = {
    "PASS": 0,
    "INFO": 0,
    "LOW": 1,
    "WARNING": 2,
    "MEDIUM": 3,
    "HIGH": 4,
    "CRITICAL": 5,
    "MANUAL CHECK": 2,
}

DATE_RE = re.compile(
    r"\b(?:\d{1,2}[\s\-/\.](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*[\s\-/\.]\d{2,4}|"
    r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*[\s\-/\.]\d{1,2}[,\s\-/\.]\d{2,4}|"
    r"\d{1,2}[\-/\.]\d{1,2}[\-/\.]\d{2,4}|"
    r"\d{4}[\-/\.]\d{1,2}[\-/\.]\d{1,2})\b",
    flags=re.IGNORECASE,
)

CERTIFICATE_NAMES = [
    "Safety Equipment Certificate",
    "Safety Radio Certificate",
    "Safety Construction Certificate",
    "Loadline Certificate",
    "International Oil Pollution Prevention Certificate",
    "IOPPC",
    "International Ballast Water Management Certificate",
    "IBWMC",
    "Safety Management Certificate",
    "SMC",
    "Document of Compliance",
    "DOC",
    "International Ship Security Certificate",
    "ISSC",
    "USCG Certificate of Compliance",
    "Minimum Safe Manning Document",
    "Civil Liability Certificate",
    "Wreck Removal Convention Certificate",
    "Certificate of Fitness",
    "Class Certificate",
]

SURVEY_FIELDS = {
    "last_dry_dock": ["date of last dry dock", "last dry dock"],
    "next_dry_dock_due": ["date next dry dock due", "next dry dock", "dry dock due"],
    "last_iws": ["date of last iws", "last iws", "in water survey"],
    "next_iws_due": ["date next iws due", "next iws", "iws due"],
    "last_special_survey": ["date of last special survey", "last special survey"],
    "next_special_survey_due": ["date next special survey due", "special survey due"],
    "last_annual_survey": ["date of last annual survey", "last annual survey", "annual survey"],
    "last_intermediate_survey": ["date of last intermediate survey", "last intermediate survey", "intermediate survey"],
    "conditions_of_class": ["conditions of class", "condition of class", "coc"],
    "memoranda_of_class": ["memoranda of class", "memorandum of class", "moc"],
    "flag_dispensation": ["flag state dispensation", "dispensation"],
}

HVPQ_HIGH_RISK_FIELDS = {
    "ship_type": ["type of ship", "if other, then specify", "product carrier", "chemical tanker"],
    "pni": ["p and i club", "p&i cover", "wreck removal"],
    "cii": ["carbon intensity indicator", "cii rating", "annual efficiency ratio", "aer"],
    "eexi": ["energy efficiency existing ship index", "eexi"],
    "eedi": ["energy efficiency design index", "eedi"],
    "eiv": ["estimated index value", "eiv"],
    "psc": ["port state control", "psc"],
    "incidents": ["pollution", "grounding", "collision", "allision", "other incidents", "type of incident"],
    "certificates": CERTIFICATE_NAMES,
    "foam": ["fixed foam", "foam", "test analysis certificate"],
    "overboard": ["overboard discharges", "blanks", "testing arrangement"],
    "cargo_pressure_test": ["cargo piping pressure test", "hydrostatically pressure test cargo"],
    "bunker_pressure_test": ["bunker piping pressure test", "hydrostatically pressure test bunker"],
    "tank_coating": ["cargo tank coating", "ballast tank coating", "coated", "coating inspection"],
    "mooring": ["mooring", "brake test", "split drum", "rendering load", "bhc"],
    "diagrams": ["mooring arrangement", "manifold arrangement", "fairlead", "chock", "bitt diagram"],
    "lifting_gear": ["lifting gear", "crane", "annual test", "five yearly"],
}


# -----------------------------
# Data models
# -----------------------------

@dataclass
class Finding:
    check_id: str
    module: str
    check: str
    status: str
    risk: str
    hvpq_value: str = ""
    piq_value: str = ""
    class_status_value: str = ""
    q88_value: str = ""
    reason: str = ""
    recommended_action: str = ""
    manual_verification_required: str = "No"
    source_location: str = ""

    @property
    def score(self) -> int:
        return RISK_SCORE.get(self.risk.upper(), 0)


@dataclass
class TextDoc:
    name: str
    kind: str
    full_text: str
    pages: List[str]


# -----------------------------
# File extraction helpers
# -----------------------------

@st.cache_data(show_spinner=False)
def read_uploaded_file(name: str, data: bytes) -> TextDoc:
    suffix = name.lower().split(".")[-1]
    if suffix == "pdf":
        return read_pdf(name, data)
    if suffix in ["xlsx", "xls"]:
        return read_excel_as_text(name, data)
    if suffix == "csv":
        return read_csv_as_text(name, data)
    if suffix in ["docx"]:
        return read_docx_as_text(name, data)
    return TextDoc(name=name, kind=suffix, full_text=data.decode("utf-8", errors="ignore"), pages=[])


def read_pdf(name: str, data: bytes) -> TextDoc:
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed. Install pymupdf.")
    pages: List[str] = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc):
            text = page.get_text("text") or ""
            pages.append(f"--- Page {i + 1} ---\n{text}")
    return TextDoc(name=name, kind="pdf", full_text="\n".join(pages), pages=pages)


def read_excel_as_text(name: str, data: bytes) -> TextDoc:
    xls = pd.ExcelFile(io.BytesIO(data))
    parts = []
    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna("")
            parts.append(f"--- Sheet: {sheet} ---")
            parts.append(df.to_csv(index=False))
        except Exception as exc:
            parts.append(f"--- Sheet: {sheet} could not be read: {exc} ---")
    return TextDoc(name=name, kind="excel", full_text="\n".join(parts), pages=parts)


def read_csv_as_text(name: str, data: bytes) -> TextDoc:
    df = pd.read_csv(io.BytesIO(data), dtype=str).fillna("")
    return TextDoc(name=name, kind="csv", full_text=df.to_csv(index=False), pages=[])


def read_docx_as_text(name: str, data: bytes) -> TextDoc:
    if Document is None:
        raise RuntimeError("python-docx is not installed. Install python-docx.")
    doc = Document(io.BytesIO(data))
    text = "\n".join(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            text += "\n" + " | ".join(cell.text for cell in row.cells)
    return TextDoc(name=name, kind="docx", full_text=text, pages=[])


# -----------------------------
# Parsing helpers
# -----------------------------

def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def lower(s: str) -> str:
    return normalize_text(s).lower()


def parse_date_safe(s: str) -> Optional[date]:
    if not s:
        return None
    s = s.strip().replace("Sept", "Sep")
    try:
        dt = dateparser.parse(s, dayfirst=True, fuzzy=True)
        if dt:
            return dt.date()
    except Exception:
        return None
    return None


def fmt_date(d: Optional[date]) -> str:
    return d.strftime("%d %b %Y") if d else ""


def dates_equal(a: Optional[date], b: Optional[date]) -> bool:
    return bool(a and b and a == b)


def month_gap_days(start: date, end: date, max_months: float) -> Tuple[bool, int, float]:
    days = (end - start).days
    months = days / 30.4375
    return months <= max_months, days, months


def find_line_after_keywords(text: str, keywords: Iterable[str], window_chars: int = 240) -> str:
    t = normalize_text(text)
    tl = t.lower()
    for kw in keywords:
        idx = tl.find(kw.lower())
        if idx >= 0:
            return t[idx : idx + window_chars]
    return ""


def find_first_date_near(text: str, keywords: Iterable[str], window_chars: int = 280) -> Tuple[str, Optional[date]]:
    snippet = find_line_after_keywords(text, keywords, window_chars=window_chars)
    if not snippet:
        return "", None
    m = DATE_RE.search(snippet)
    return snippet, parse_date_safe(m.group(0)) if m else None


def find_yes_no_near(text: str, keywords: Iterable[str], window_chars: int = 280) -> Tuple[str, str]:
    snippet = find_line_after_keywords(text, keywords, window_chars=window_chars)
    if not snippet:
        return "", ""
    # Prefer explicit Yes/No following the keyword block.
    m = re.search(r"\b(Yes|No|Not applicable|NA|N/A)\b", snippet, flags=re.IGNORECASE)
    return snippet, m.group(1).title() if m else ""


def extract_value_after_label(text: str, labels: Iterable[str], max_chars: int = 120) -> Tuple[str, str]:
    t = normalize_text(text)
    tl = t.lower()
    for label in labels:
        idx = tl.find(label.lower())
        if idx >= 0:
            snippet = t[idx : idx + max_chars]
            tail = t[idx + len(label) : idx + max_chars]
            # Stop at likely next numbered question if present.
            tail = re.split(r"\s\d+\.\d+(?:\.\d+)?\s", tail)[0]
            return snippet, normalize_text(tail)
    return "", ""


def fuzzy_contains(text: str, phrase: str, threshold: int = 88) -> bool:
    txt = lower(text)
    ph = lower(phrase)
    if ph in txt:
        return True
    # Only fuzzy on short phrase chunks to avoid heavy computation.
    for chunk in re.split(r"[\n\.;:]", text):
        if fuzz.partial_ratio(lower(chunk), ph) >= threshold:
            return True
    return False


# -----------------------------
# Document-specific extraction
# -----------------------------

def extract_hvpq_fields(doc: Optional[TextDoc]) -> Dict[str, Any]:
    if doc is None:
        return {}
    text = doc.full_text
    fields: Dict[str, Any] = {"doc_name": doc.name, "_full_text": text}

    # Identity
    _, fields["vessel_name"] = extract_value_after_label(text, ["Name of ship", "Vessel Particulars Questionnaire for"], 90)
    if "Vessel Particulars Questionnaire for" in text and not fields.get("vessel_name"):
        m = re.search(r"Vessel Particulars Questionnaire for\s+(.+?)\s+IMO", text, flags=re.I)
        if m:
            fields["vessel_name"] = normalize_text(m.group(1))
    m = re.search(r"(?:IMO/LR Number|LR/IMO number|IMO:)\s*(\d{7})", text, flags=re.I)
    fields["imo"] = m.group(1) if m else ""
    _, doc_date = find_first_date_near(text, ["Date this HVPQ document completed", "Harmonised Vessel Particulars Questionnaire"])
    fields["doc_date"] = doc_date

    # Vessel type
    _, fields["ship_type"] = extract_value_after_label(text, ["What is the type of ship", "Type of ship"], 160)
    _, fields["ship_type_other"] = extract_value_after_label(text, ["If other, then specify"], 120)

    # Environmental
    for key, labels in {
        "cii_rating": ["provide CII rating", "CII rating"],
        "cii_verified_by": ["Is the CII rating verified by", "CII rating verified"],
        "eexi_rating": ["provide EEXI rating", "EEXI rating"],
        "eedi_rating": ["provide EEDI rating", "EEDI rating"],
        "eiv_rating": ["provide EIV rating", "EIV rating"],
    }.items():
        _, val = extract_value_after_label(text, labels, 140)
        fields[key] = val

    # Class/survey fields
    for key, labels in SURVEY_FIELDS.items():
        snippet, d = find_first_date_near(text, labels)
        fields[key] = d
        fields[key + "_snippet"] = snippet
        if key in ["conditions_of_class", "memoranda_of_class", "flag_dispensation"]:
            snippet2, yn = find_yes_no_near(text, labels)
            fields[key + "_yn"] = yn
            fields[key + "_snippet"] = snippet2 or snippet

    # PSC
    snippet, d = find_first_date_near(text, ["Date of last Port State Control Inspection", "last Port State Control Inspection"])
    fields["last_psc_date"] = d
    fields["last_psc_snippet"] = snippet
    _, fields["last_psc_port"] = extract_value_after_label(text, ["Port of last Port State Control Inspection"], 120)
    snippet, yn = find_yes_no_near(text, ["detained during the last 36 months"])
    fields["detained_36m"] = yn
    fields["detained_36m_snippet"] = snippet

    # Incidents
    snippet, yn = find_yes_no_near(text, ["pollution, grounding, collision or allision incident during the past 12 months"])
    fields["pollution_grounding_collision_allision_12m"] = yn
    fields["pollution_grounding_collision_allision_snippet"] = snippet
    snippet, yn = find_yes_no_near(text, ["any other incidents during the past 12 months"])
    fields["other_incidents_12m"] = yn
    fields["other_incidents_snippet"] = snippet
    fields["incident_detail_text"] = find_line_after_keywords(text, ["If yes, provide details (see table)", "Type of Incident"], 500)

    # Certificates: generic extraction of lines around certificate names
    fields["certificates"] = extract_certificate_dates_generic(text)

    # Targeted high-observation areas: generic snippets/dates from HVPQ.
    targeted_keywords = {
        "mooring_block": ["mooring", "brake test", "brake holding", "BHC", "rendering load", "split drum", "rope", "tail", "end-for-end"],
        "tank_block": ["cargo tank coating", "ballast tank coating", "void", "coating inspection", "Frequency of Inspections"],
        "piping_pressure_block": ["cargo piping pressure tests", "bunker piping pressure tests", "hydrostatically pressure test"],
        "pollution_block": ["overboard discharges", "sea valves", "seachest", "scupper", "spill containment"],
        "fire_foam_block": ["fixed foam", "foam", "Test Analysis Certificate", "firefighting system"],
        "lifting_block": ["crane", "lifting", "SWL", "annual test", "five year"],
        "diagram_block": ["mooring arrangement", "manifold arrangement", "fairlead", "chock", "bitt diagram", "bow mooring"],
    }
    for k, kws in targeted_keywords.items():
        fields[k] = find_line_after_keywords(text, kws, 1800)
        fields[k + "_dates"] = extract_dates_from_block(text, kws, max_dates=8, window_chars=1800)
    return fields


def extract_piq_fields(doc: Optional[TextDoc]) -> Dict[str, Any]:
    if doc is None:
        return {}
    text = doc.full_text
    fields: Dict[str, Any] = {"doc_name": doc.name, "_full_text": text}
    _, fields["vessel_name"] = extract_value_after_label(text, ["Vessel Name"], 90)
    snippet, fields["doc_date"] = find_first_date_near(text, ["Date"])
    _, fields["vessel_type"] = extract_value_after_label(text, ["Vessel Type"], 160)

    # Superintendent visits - extract dates near the block.
    fields["technical_superintendent_dates"] = extract_dates_from_block(
        text,
        ["Technical Superintendent inspection completed", "Technical Superintendent"],
        max_dates=8,
        window_chars=850,
    )
    fields["marine_superintendent_dates"] = extract_dates_from_block(
        text,
        ["Marine Superintendent inspection completed", "Marine Superintendent"],
        max_dates=8,
        window_chars=650,
    )

    # PSC table block
    fields["psc_dates"] = extract_dates_from_block(text, ["last three Port State Control", "PSC inspection"], max_dates=6, window_chars=1000)
    fields["psc_block"] = find_line_after_keywords(text, ["last three Port State Control", "PSC inspection"], 1200)

    # Tank inspections
    for key, labels in {
        "cargo_tank_oldest_inspection": ["oldest inspection report for all cargo and slop tanks"],
        "ballast_tank_oldest_inspection": ["oldest inspection report for all ballast tanks"],
        "void_space_oldest_inspection": ["oldest inspection report for all void space", "void spaces"],
    }.items():
        snippet, d = find_first_date_near(text, labels)
        fields[key] = d
        fields[key + "_snippet"] = snippet

    # MOC / retrofit
    _, fields["structural_changes"] = find_yes_no_near(text, ["structural changes been made"])
    snippet, yn = find_yes_no_near(text, ["equipment been retrofitted"])
    fields["equipment_retrofitted"] = yn
    fields["equipment_retrofitted_snippet"] = snippet
    _, fields["equipment_replaced"] = find_yes_no_near(text, ["non like-for-like", "non-like-for-like"])
    _, fields["equipment_decommissioned"] = find_yes_no_near(text, ["equipment been decommissioned"])

    # Incident declarations from PIQ can vary. Gather blocks and yes/no around incident terms.
    incident_keywords = [
        "incident", "lost time", "injury", "pollution", "grounding", "collision", "allision",
        "blackout", "machinery failure", "mooring", "rope parting", "fire", "explosion",
    ]
    fields["incident_blocks"] = collect_keyword_blocks(text, incident_keywords, max_blocks=20, window_chars=320)
    fields["incident_yes_count"] = count_yes_near_blocks(fields["incident_blocks"])
    fields["incident_no_count"] = count_no_near_blocks(fields["incident_blocks"])

    # PIQ targeted snippets useful for ship verification checklists.
    piq_targeted_keywords = {
        "tank_inspection_block": ["required frequency of inspection for cargo tanks", "ballast tanks", "void spaces", "oldest inspection report"],
        "moc_block": ["structural changes", "retrofitted", "equipment replaced", "decommissioned"],
        "management_oversight_block": ["Technical Superintendent", "Marine Superintendent", "full inspection"],
    }
    for k, kws in piq_targeted_keywords.items():
        fields[k] = find_line_after_keywords(text, kws, 1600)
        fields[k + "_dates"] = extract_dates_from_block(text, kws, max_dates=10, window_chars=1600)
    return fields


def extract_reference_fields(doc: Optional[TextDoc], label: str) -> Dict[str, Any]:
    """Generic extractor for Class Status / Q88 / certificate packs."""
    if doc is None:
        return {}
    text = doc.full_text
    fields: Dict[str, Any] = {"doc_name": doc.name, "label": label, "_full_text": text}
    for key, labels in SURVEY_FIELDS.items():
        snippet, d = find_first_date_near(text, labels, window_chars=360)
        fields[key] = d
        fields[key + "_snippet"] = snippet
        if key in ["conditions_of_class", "memoranda_of_class", "flag_dispensation"]:
            _, yn = find_yes_no_near(text, labels, window_chars=360)
            fields[key + "_yn"] = yn
    fields["certificates"] = extract_certificate_dates_generic(text)
    fields["psc_dates"] = extract_dates_from_block(text, ["Port State Control", "PSC"], max_dates=10, window_chars=1600)
    return fields


def extract_dates_from_block(text: str, keywords: Iterable[str], max_dates: int = 10, window_chars: int = 900) -> List[date]:
    block = find_line_after_keywords(text, keywords, window_chars=window_chars)
    dates = []
    for m in DATE_RE.finditer(block):
        d = parse_date_safe(m.group(0))
        if d and d not in dates:
            dates.append(d)
        if len(dates) >= max_dates:
            break
    return dates


def collect_keyword_blocks(text: str, keywords: Iterable[str], max_blocks: int = 20, window_chars: int = 260) -> List[str]:
    blocks = []
    tl = text.lower()
    for kw in keywords:
        for m in re.finditer(re.escape(kw.lower()), tl):
            start = max(0, m.start() - 80)
            end = min(len(text), m.end() + window_chars)
            block = normalize_text(text[start:end])
            if block and block not in blocks:
                blocks.append(block)
            if len(blocks) >= max_blocks:
                return blocks
    return blocks


def count_yes_near_blocks(blocks: List[str]) -> int:
    return sum(1 for b in blocks if re.search(r"\bYes\b", b, flags=re.I))


def count_no_near_blocks(blocks: List[str]) -> int:
    return sum(1 for b in blocks if re.search(r"\bNo\b", b, flags=re.I))


def extract_certificate_dates_generic(text: str) -> Dict[str, Dict[str, Any]]:
    results: Dict[str, Dict[str, Any]] = {}
    norm = normalize_text(text)
    nlow = norm.lower()
    for cert in CERTIFICATE_NAMES:
        idx = nlow.find(cert.lower())
        if idx < 0:
            continue
        snippet = norm[idx : idx + 420]
        dates = []
        for m in DATE_RE.finditer(snippet):
            d = parse_date_safe(m.group(0))
            if d and d not in dates:
                dates.append(d)
        results[cert] = {"snippet": snippet, "dates": dates}
    return results


# -----------------------------
# Rules
# -----------------------------

def run_all_rules(
    hvpq: Dict[str, Any],
    piq: Dict[str, Any],
    class_status: Dict[str, Any],
    q88: Dict[str, Any],
    settings: Dict[str, Any],
    obs_docs: List[TextDoc],
) -> List[Finding]:
    findings: List[Finding] = []
    findings.extend(rule_identity(hvpq, piq))
    findings.extend(rule_piq_vs_hvpq_general(hvpq, piq))
    findings.extend(rule_incident_declarations(hvpq, piq))
    findings.extend(rule_superintendent_visits(piq, hvpq, settings))
    findings.extend(rule_psc_cross_check(hvpq, piq, class_status, q88))
    findings.extend(rule_survey_and_certificate_due(hvpq, settings))
    findings.extend(rule_class_status_cross_check(hvpq, class_status, source_label="Class Status"))
    findings.extend(rule_class_status_cross_check(hvpq, q88, source_label="Q88"))
    findings.extend(rule_cii(hvpq, q88))
    findings.extend(rule_moc_retrofit(hvpq, piq))
    findings.extend(rule_targeted_operational_checks(hvpq, piq, class_status, q88, settings))
    findings.extend(rule_observation_library_scan(hvpq, piq, obs_docs))
    return findings


def rule_identity(hvpq: Dict[str, Any], piq: Dict[str, Any]) -> List[Finding]:
    out = []
    hv_name = normalize_text(hvpq.get("vessel_name", ""))
    piq_name = normalize_text(piq.get("vessel_name", ""))
    if hv_name and piq_name:
        score = fuzz.token_sort_ratio(hv_name, piq_name)
        if score >= 90:
            out.append(Finding("ID-001", "Identity", "Vessel name HVPQ vs PIQ", "PASS", "PASS", hv_name, piq_name, reason="Vessel names appear to match."))
        else:
            out.append(Finding("ID-001", "Identity", "Vessel name HVPQ vs PIQ", "FAIL", "CRITICAL", hv_name, piq_name, reason="Vessel names do not match.", recommended_action="Confirm correct HVPQ/PIQ documents have been uploaded."))
    else:
        out.append(Finding("ID-001", "Identity", "Vessel name HVPQ vs PIQ", "MANUAL CHECK", "HIGH", hv_name, piq_name, reason="Vessel name could not be extracted from one or both documents.", recommended_action="Manually verify document identity."))

    hv_date = hvpq.get("doc_date")
    piq_date = piq.get("doc_date")
    if hv_date and piq_date:
        delta = abs((hv_date - piq_date).days)
        risk = "PASS" if delta <= 7 else "WARNING"
        out.append(Finding("ID-002", "Identity", "Document date alignment", "PASS" if delta <= 7 else "WARNING", risk, fmt_date(hv_date), fmt_date(piq_date), reason=f"HVPQ/PIQ document date difference is {delta} days."))
    else:
        out.append(Finding("ID-002", "Identity", "Document date alignment", "MANUAL CHECK", "LOW", fmt_date(hv_date), fmt_date(piq_date), reason="Could not extract one or both document dates."))
    return out


def rule_piq_vs_hvpq_general(hvpq: Dict[str, Any], piq: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    hv_type = " ".join([normalize_text(hvpq.get("ship_type", "")), normalize_text(hvpq.get("ship_type_other", ""))]).strip()
    piq_type = normalize_text(piq.get("vessel_type", ""))
    if hv_type or piq_type:
        if hv_type and piq_type and fuzz.partial_ratio(hv_type.lower(), piq_type.lower()) >= 65:
            out.append(Finding("GEN-001", "PIQ vs HVPQ", "Vessel type consistency", "PASS", "PASS", hv_type, piq_type, reason="Vessel type broadly aligns."))
        else:
            out.append(Finding("GEN-001", "PIQ vs HVPQ", "Vessel type consistency", "MANUAL CHECK", "MEDIUM", hv_type, piq_type, reason="Vessel type wording differs or could not be matched. Product Carrier vs Products/Chemical Tanker may be acceptable but should be confirmed against IOPP/CoF/Annex II carriage.", recommended_action="Ask vessel/office to confirm correct vessel type wording and Annex II carriage intention."))

    # Tank inspection dates in PIQ against HVPQ tank coating text presence.
    for key, label in [
        ("cargo_tank_oldest_inspection", "Cargo/slop tank inspection sequence"),
        ("ballast_tank_oldest_inspection", "Ballast tank inspection sequence"),
        ("void_space_oldest_inspection", "Void space inspection sequence"),
    ]:
        d = piq.get(key)
        if d:
            out.append(Finding(f"GEN-{key}", "PIQ General Accuracy", label, "INFO", "INFO", "", fmt_date(d), reason="PIQ date extracted. Confirm supporting tank/void inspection reports are available onboard.", recommended_action="Ship to verify the oldest report date in the current sequence and evidence availability.", manual_verification_required="Yes"))
        else:
            out.append(Finding(f"GEN-{key}", "PIQ General Accuracy", label, "MANUAL CHECK", "MEDIUM", "", "Not extracted", reason="Could not extract PIQ inspection sequence date.", recommended_action="Ship to verify field and supporting evidence."))
    return out


def rule_incident_declarations(hvpq: Dict[str, Any], piq: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    hv_poll = normalize_text(hvpq.get("pollution_grounding_collision_allision_12m", ""))
    hv_other = normalize_text(hvpq.get("other_incidents_12m", ""))
    hv_detail = normalize_text(hvpq.get("incident_detail_text", ""))
    piq_yes = piq.get("incident_yes_count", 0)
    piq_no = piq.get("incident_no_count", 0)

    # User's instruction: no JiBe list. If no incidents entered, flag no incidents entered.
    if hv_poll.lower() in ["no", "n/a", "na", ""] and hv_other.lower() in ["no", "n/a", "na", ""]:
        out.append(Finding(
            "INC-001", "Incident Declaration", "HVPQ incident declaration blank/no",
            "MANUAL CHECK", "HIGH",
            hvpq_value=f"Pollution/grounding/collision/allision: {hv_poll or 'Not extracted'}; Other incidents: {hv_other or 'Not extracted'}",
            piq_value="",
            reason="HVPQ appears to have no incidents entered or incident fields could not be extracted. This should be positively confirmed, not assumed correct.",
            recommended_action="Send to vessel/office: confirm there were no reportable incidents, injuries, machinery failures, pollution, mooring, navigation or equipment incidents during the HVPQ/PIQ lookback period.",
            manual_verification_required="Yes",
        ))
    else:
        out.append(Finding(
            "INC-001", "Incident Declaration", "HVPQ incident declaration present",
            "INFO", "INFO",
            hvpq_value=f"Pollution/grounding/collision/allision: {hv_poll}; Other incidents: {hv_other}; Details: {hv_detail[:220]}",
            reason="HVPQ has incident declaration data. Verify that PIQ incident questions and details are consistent.",
            recommended_action="Ship/office to confirm all incidents in HVPQ are correctly reflected in PIQ where applicable.",
            manual_verification_required="Yes",
        ))

    # Mismatch logic from document declarations only.
    if hv_other.lower() == "yes" and piq_yes == 0:
        out.append(Finding(
            "INC-002", "Incident Declaration", "HVPQ incident yes vs PIQ incident no/blank",
            "FAIL", "HIGH",
            hvpq_value=f"Other incidents: {hv_other}; details: {hv_detail[:220]}",
            piq_value=f"PIQ incident yes count: {piq_yes}, no count: {piq_no}",
            reason="HVPQ indicates other incidents but PIQ incident-related positive declarations were not detected.",
            recommended_action="Check PIQ incident section manually and update relevant PIQ fields if required.",
            manual_verification_required="Yes",
        ))
    if hv_poll.lower() == "yes" and piq_yes == 0:
        out.append(Finding(
            "INC-003", "Incident Declaration", "Major incident declaration mismatch",
            "FAIL", "CRITICAL",
            hvpq_value=f"Pollution/grounding/collision/allision: {hv_poll}",
            piq_value=f"PIQ incident yes count: {piq_yes}",
            reason="HVPQ indicates pollution/grounding/collision/allision but PIQ positive declaration was not detected.",
            recommended_action="Urgently verify PIQ and supporting incident evidence before inspection.",
            manual_verification_required="Yes",
        ))
    if piq_yes == 0:
        out.append(Finding(
            "INC-004", "Incident Declaration", "PIQ incident positive entries not detected",
            "MANUAL CHECK", "MEDIUM",
            hvpq_value=f"HVPQ other incidents: {hv_other}",
            piq_value="No clear PIQ incident Yes detected from extracted text",
            reason="PIQ incident sections vary in wording. No positive incident declaration was detected by text parsing.",
            recommended_action="Ship to verify PIQ incident questions one by one, especially injury, machinery, mooring, navigation, pollution and security-related questions.",
            manual_verification_required="Yes",
        ))
    return out


def rule_superintendent_visits(piq: Dict[str, Any], hvpq: Dict[str, Any], settings: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    ref_date = piq.get("doc_date") or hvpq.get("doc_date") or date.today()
    for label, key, max_months in [
        ("Technical Superintendent", "technical_superintendent_dates", settings.get("ts_gap_months", DEFAULT_TS_GAP_MONTHS)),
        ("Marine Superintendent", "marine_superintendent_dates", settings.get("ms_gap_months", DEFAULT_MS_GAP_MONTHS)),
    ]:
        dates = sorted(piq.get(key, []))
        if not dates:
            out.append(Finding(f"SUP-{key}-000", "Management Oversight", f"{label} visit dates", "MANUAL CHECK", "HIGH", piq_value="No dates extracted", reason=f"{label} visit dates could not be extracted from PIQ.", recommended_action="Verify PIQ management oversight table and supporting visit reports.", manual_verification_required="Yes"))
            continue
        # Check gaps between successive visit end/start dates approximately. Dates extracted include from/to, so use unique sorted dates.
        all_dates = dates + [ref_date]
        for i in range(len(all_dates) - 1):
            ok, days, months = month_gap_days(all_dates[i], all_dates[i + 1], float(max_months))
            out.append(Finding(
                f"SUP-{key}-{i+1:03d}", "Management Oversight", f"{label} gap check",
                "PASS" if ok else "FAIL", "PASS" if ok else "CRITICAL",
                piq_value=f"{fmt_date(all_dates[i])} → {fmt_date(all_dates[i + 1])}",
                reason=f"Gap is {days} days / {months:.2f} months. Limit is {max_months:.1f} months strict.",
                recommended_action="No action." if ok else f"Explain gap and arrange/record {label} full inspection as required.",
                manual_verification_required="No" if ok else "Yes",
            ))
    return out


def rule_psc_cross_check(hvpq: Dict[str, Any], piq: Dict[str, Any], class_status: Dict[str, Any], q88: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    hv_date = hvpq.get("last_psc_date")
    piq_dates = piq.get("psc_dates", [])
    piq_last = max(piq_dates) if piq_dates else None
    if hv_date and piq_last:
        if hv_date == piq_last:
            out.append(Finding("PSC-001", "PSC", "Last PSC date HVPQ vs PIQ", "PASS", "PASS", fmt_date(hv_date), fmt_date(piq_last), reason="Last PSC dates match."))
        else:
            out.append(Finding("PSC-001", "PSC", "Last PSC date HVPQ vs PIQ", "FAIL", "HIGH", fmt_date(hv_date), fmt_date(piq_last), reason="Last PSC dates differ.", recommended_action="Verify latest PSC inspection date, port, MOU, deficiencies and detention status."))
    else:
        out.append(Finding("PSC-001", "PSC", "Last PSC date HVPQ vs PIQ", "MANUAL CHECK", "MEDIUM", fmt_date(hv_date), fmt_date(piq_last), reason="Could not extract PSC date from one or both documents.", recommended_action="Ship/office to confirm latest PSC data."))

    hv_det = normalize_text(hvpq.get("detained_36m", ""))
    if not hv_det:
        out.append(Finding("PSC-002", "PSC", "Detention status in HVPQ", "MANUAL CHECK", "HIGH", hvpq_value="Not extracted", reason="Detention status for last 36 months was not extracted.", recommended_action="Confirm HVPQ detention status and PSC database entries."))
    return out


def rule_survey_and_certificate_due(hvpq: Dict[str, Any], settings: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    ref_date = hvpq.get("doc_date") or date.today()
    survey_warn = int(settings.get("survey_warning_days", DEFAULT_SURVEY_WARNING_DAYS))
    cert_warn = int(settings.get("cert_warning_days", DEFAULT_CERT_WARNING_DAYS))

    for key in ["next_dry_dock_due", "next_iws_due", "next_special_survey_due"]:
        d = hvpq.get(key)
        label = key.replace("_", " ").title()
        if not d:
            out.append(Finding(f"DUE-{key}", "HVPQ Dates", label, "MANUAL CHECK", "MEDIUM", hvpq_value="Not extracted", reason="Due date not extracted from HVPQ.", recommended_action="Verify against Class Status."))
            continue
        days = (d - ref_date).days
        if days < 0:
            out.append(Finding(f"DUE-{key}", "HVPQ Dates", label, "FAIL", "CRITICAL", hvpq_value=fmt_date(d), reason=f"Due date appears overdue by {-days} days as of {fmt_date(ref_date)}.", recommended_action="Verify immediately against Class Status and update HVPQ."))
        elif days <= survey_warn:
            out.append(Finding(f"DUE-{key}", "HVPQ Dates", label, "WARNING", "MEDIUM", hvpq_value=fmt_date(d), reason=f"Due date is within {survey_warn} days.", recommended_action="Confirm planning and HVPQ accuracy."))
        else:
            out.append(Finding(f"DUE-{key}", "HVPQ Dates", label, "PASS", "PASS", hvpq_value=fmt_date(d), reason=f"Due date is {days} days away."))

    # Certificate expiry dates from HVPQ table, if extracted.
    certs = hvpq.get("certificates", {})
    for cert, info in certs.items():
        dates = info.get("dates", [])
        if len(dates) < 2:
            continue
        # For certificate lines, expiry is often second date; this is generic and should be manually verified.
        expiry = dates[1]
        days = (expiry - ref_date).days
        if days < 0:
            risk, status = "CRITICAL", "FAIL"
        elif days <= cert_warn:
            risk, status = "HIGH", "WARNING"
        else:
            risk, status = "INFO", "INFO"
        if risk != "INFO":
            out.append(Finding(
                f"CERT-DUE-{cert[:20]}", "Certificate Due", cert, status, risk,
                hvpq_value=f"Possible expiry: {fmt_date(expiry)}",
                reason=f"Generic extraction found certificate date within/over warning window. Days to expiry: {days}.",
                recommended_action="Confirm exact certificate issue/expiry/endorsement dates against certificate or Class Status.",
                manual_verification_required="Yes",
                source_location=info.get("snippet", "")[:250],
            ))
    return out


def rule_class_status_cross_check(hvpq: Dict[str, Any], ref: Dict[str, Any], source_label: str) -> List[Finding]:
    out: List[Finding] = []
    if not ref:
        out.append(Finding(f"{source_label}-000", source_label, f"{source_label} uploaded", "INFO", "INFO", reason=f"No {source_label} file uploaded. Cross-check skipped."))
        return out

    for key, labels in SURVEY_FIELDS.items():
        if key in ["conditions_of_class", "memoranda_of_class", "flag_dispensation"]:
            hv = normalize_text(hvpq.get(key + "_yn", ""))
            rv = normalize_text(ref.get(key + "_yn", ""))
            if hv and rv and hv.lower() != rv.lower():
                out.append(Finding(f"{source_label}-{key}", source_label, key.replace("_", " ").title(), "FAIL", "HIGH", hvpq_value=hv, class_status_value=rv if source_label == "Class Status" else "", q88_value=rv if source_label == "Q88" else "", reason=f"HVPQ and {source_label} appear to differ for {key}.", recommended_action=f"Verify {key} in latest {source_label} and update HVPQ if required.", manual_verification_required="Yes"))
            elif hv or rv:
                out.append(Finding(f"{source_label}-{key}", source_label, key.replace("_", " ").title(), "INFO", "INFO", hvpq_value=hv, class_status_value=rv if source_label == "Class Status" else "", q88_value=rv if source_label == "Q88" else "", reason=f"{key} values extracted for reference. Confirm if parsing is uncertain."))
            continue

        hvd = hvpq.get(key)
        rd = ref.get(key)
        if hvd and rd:
            if hvd == rd:
                out.append(Finding(f"{source_label}-{key}", source_label, key.replace("_", " ").title(), "PASS", "PASS", hvpq_value=fmt_date(hvd), class_status_value=fmt_date(rd) if source_label == "Class Status" else "", q88_value=fmt_date(rd) if source_label == "Q88" else "", reason=f"HVPQ date matches {source_label}."))
            else:
                out.append(Finding(f"{source_label}-{key}", source_label, key.replace("_", " ").title(), "FAIL", "HIGH", hvpq_value=fmt_date(hvd), class_status_value=fmt_date(rd) if source_label == "Class Status" else "", q88_value=fmt_date(rd) if source_label == "Q88" else "", reason=f"HVPQ date does not match {source_label} extracted date.", recommended_action=f"Verify exact value in latest {source_label}; update HVPQ if wrong.", manual_verification_required="Yes", source_location=normalize_text(ref.get(key + "_snippet", ""))[:250]))
        elif hvd and not rd:
            out.append(Finding(f"{source_label}-{key}", source_label, key.replace("_", " ").title(), "MANUAL CHECK", "LOW", hvpq_value=fmt_date(hvd), reason=f"HVPQ date exists but matching {source_label} date was not extracted. Class status formats vary; this is not automatically a defect.", recommended_action=f"Manual check against {source_label}.", manual_verification_required="Yes"))
    # Certificates generic comparison by name.
    hv_certs = hvpq.get("certificates", {})
    ref_certs = ref.get("certificates", {})
    for cert, hv_info in hv_certs.items():
        if cert not in ref_certs:
            continue
        hv_dates = hv_info.get("dates", [])
        ref_dates = ref_certs[cert].get("dates", [])
        if hv_dates and ref_dates:
            if set(hv_dates[:3]).intersection(ref_dates[:4]):
                out.append(Finding(f"{source_label}-CERT-{cert[:18]}", source_label, f"{cert} date cross-check", "INFO", "INFO", hvpq_value=", ".join(fmt_date(d) for d in hv_dates[:4]), class_status_value=", ".join(fmt_date(d) for d in ref_dates[:4]) if source_label == "Class Status" else "", q88_value=", ".join(fmt_date(d) for d in ref_dates[:4]) if source_label == "Q88" else "", reason=f"Some certificate dates overlap between HVPQ and {source_label}. Confirm exact issue/expiry/endorsement columns manually."))
            else:
                out.append(Finding(f"{source_label}-CERT-{cert[:18]}", source_label, f"{cert} date cross-check", "MANUAL CHECK", "MEDIUM", hvpq_value=", ".join(fmt_date(d) for d in hv_dates[:4]), class_status_value=", ".join(fmt_date(d) for d in ref_dates[:4]) if source_label == "Class Status" else "", q88_value=", ".join(fmt_date(d) for d in ref_dates[:4]) if source_label == "Q88" else "", reason=f"Certificate date sets do not visibly overlap between HVPQ and {source_label}. Generic extraction may be imperfect.", recommended_action=f"Confirm {cert} dates against original certificate/{source_label}.", manual_verification_required="Yes"))
    return out


def rule_cii(hvpq: Dict[str, Any], q88: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    cii = normalize_text(hvpq.get("cii_rating", ""))
    verified = normalize_text(hvpq.get("cii_verified_by", ""))
    if not cii or cii.upper() in ["NA", "N/A", "NO"]:
        out.append(Finding("CII-001", "Environmental", "CII rating declared", "FAIL", "HIGH", hvpq_value=cii or "Blank/not extracted", reason="CII rating is blank/not extracted or appears not declared.", recommended_action="Verify latest CII/AER evidence and update HVPQ."))
    else:
        risk = "WARNING" if "owner" in verified.lower() else "INFO"
        out.append(Finding("CII-001", "Environmental", "CII rating declared", "INFO", risk, hvpq_value=f"CII: {cii}; verified by: {verified}", reason="CII data extracted from HVPQ. If verified by Owner only, confirm latest class/recognized evidence if required by vetting party.", recommended_action="Ship/office to attach latest CII/AER evidence for verification.", manual_verification_required="Yes"))
    return out


def rule_moc_retrofit(hvpq: Dict[str, Any], piq: Dict[str, Any]) -> List[Finding]:
    out: List[Finding] = []
    retrofit = normalize_text(piq.get("equipment_retrofitted", ""))
    retrofit_snip = normalize_text(piq.get("equipment_retrofitted_snippet", ""))
    hv_text = " ".join([str(v) for v in hvpq.values() if isinstance(v, str)])
    if retrofit.lower() == "yes":
        if any(term in retrofit_snip.lower() for term in ["egcs", "scrubber", "sox"]):
            if re.search(r"\b(EGCS|SOx|scrubber)\b", hv_text, flags=re.I):
                out.append(Finding("MOC-001", "MOC / Retrofit", "EGCS retrofit reflected in HVPQ", "PASS", "PASS", hvpq_value="EGCS/SOx text detected", piq_value=retrofit_snip[:220], reason="PIQ retrofit item appears reflected in HVPQ text/class notation."))
            else:
                out.append(Finding("MOC-001", "MOC / Retrofit", "EGCS retrofit reflected in HVPQ", "MANUAL CHECK", "HIGH", hvpq_value="EGCS/SOx not detected", piq_value=retrofit_snip[:220], reason="PIQ mentions EGCS/SOx retrofit but related HVPQ text was not detected.", recommended_action="Verify class notation/certificates/equipment sections and update HVPQ if required.", manual_verification_required="Yes"))
        out.append(Finding("MOC-002", "MOC / Retrofit", "Retrofitted equipment declaration", "INFO", "MEDIUM", piq_value=retrofit_snip[:250], reason="PIQ states equipment was retrofitted. HVPQ, class status and relevant certificates should be checked for consistency.", recommended_action="Ship/office to confirm all retrofits are correctly reflected in HVPQ and certificates.", manual_verification_required="Yes"))
    return out



def rule_targeted_operational_checks(hvpq: Dict[str, Any], piq: Dict[str, Any], class_status: Dict[str, Any], q88: Dict[str, Any], settings: Dict[str, Any]) -> List[Finding]:
    """
    Detailed targeted checks based on recurring HVPQ/PIQ observations.
    These checks are intentionally objective: if a value cannot be reliably extracted,
    the output is a ship-verification item rather than an assumed defect.
    """
    out: List[Finding] = []
    ref_date = piq.get("doc_date") or hvpq.get("doc_date") or date.today()

    def has_text(source: Dict[str, Any], key: str) -> bool:
        return bool(normalize_text(source.get(key, "")))

    def snippet(source: Dict[str, Any], key: str, n: int = 300) -> str:
        return normalize_text(source.get(key, ""))[:n]

    # 1. Tank / void inspection objective due checks from PIQ.
    tank_items = [
        ("TANK-001", "Cargo/slop tank inspection sequence", "cargo_tank_oldest_inspection", 12),
        ("TANK-002", "Ballast tank inspection sequence", "ballast_tank_oldest_inspection", 12),
        ("TANK-003", "Void space inspection sequence", "void_space_oldest_inspection", 12),
    ]
    warn_days = int(settings.get("tank_warning_days", 60))
    for cid, label, key, freq_months in tank_items:
        d = piq.get(key)
        if d:
            due = d + relativedelta(months=freq_months)
            days_left = (due - ref_date).days
            if days_left < 0:
                status, risk = "FAIL", "HIGH"
                reason = f"PIQ oldest inspection date is {fmt_date(d)}; next due calculated as {fmt_date(due)}, overdue by {-days_left} days."
            elif days_left <= warn_days:
                status, risk = "WARNING", "MEDIUM"
                reason = f"PIQ oldest inspection date is {fmt_date(d)}; next due calculated as {fmt_date(due)}, due in {days_left} days."
            else:
                status, risk = "INFO", "INFO"
                reason = f"PIQ oldest inspection date is {fmt_date(d)}; next due calculated as {fmt_date(due)}."
            out.append(Finding(cid, "Tank / Structural", label, status, risk, piq_value=f"Oldest: {fmt_date(d)}; calculated due: {fmt_date(due)}", reason=reason, recommended_action="Ship to verify actual tank/void inspection report sequence and confirm all applicable spaces are covered.", manual_verification_required="Yes"))
        else:
            out.append(Finding(cid, "Tank / Structural", label, "MANUAL CHECK", "MEDIUM", piq_value="Not extracted", reason="PIQ oldest inspection date could not be extracted.", recommended_action="Ship to confirm oldest inspection report date and inspection frequency for this category.", manual_verification_required="Yes"))

    # HVPQ tank coating/table presence check.
    if has_text(hvpq, "tank_block"):
        out.append(Finding("TANK-004", "Tank / Structural", "HVPQ tank coating / inspection table presence", "INFO", "INFO", hvpq_value=snippet(hvpq, "tank_block"), reason="HVPQ tank coating/inspection text detected. This should be checked for full tank listing and date accuracy.", recommended_action="Ship to verify all cargo, slop, ballast and applicable void spaces are individually listed with coating condition and last inspection dates.", manual_verification_required="Yes"))
    else:
        out.append(Finding("TANK-004", "Tank / Structural", "HVPQ tank coating / inspection table presence", "MANUAL CHECK", "HIGH", hvpq_value="Not detected", reason="HVPQ tank coating/inspection section was not detected by text extraction.", recommended_action="Verify HVPQ tank coating/inspection entries manually.", manual_verification_required="Yes"))

    # 2. Mooring / ropes / brake testing. These are often not reliably extractable from all HVPQs, so use targeted objective prompts.
    moor_text = (snippet(hvpq, "mooring_block", 900) + " " + hvpq.get("_full_text", "")[:0])
    moor_dates = hvpq.get("mooring_block_dates", [])
    mooring_checks = [
        ("MOOR-001", "Mooring winch brake test date", ["brake test", "brake", "BHC", "brake holding"], "Confirm last brake test date, brake holding capacity and test certificate/report. Update HVPQ if date or value is wrong."),
        ("MOOR-002", "Brake holding capacity / rendering load", ["bhc", "brake holding", "rendering load"], "Confirm BHC/rendering load values against mooring equipment records and brake test certificates."),
        ("MOOR-003", "Split drum declaration", ["split drum", "split"], "Confirm whether winches are split drum/non-split drum and HVPQ declaration is correct."),
        ("MOOR-004", "Mooring rope / tail particulars", ["rope", "tail", "mbs", "ldbf", "material"], "Confirm rope/tail material, diameter, MBL/LDBF, certificate dates and compatibility with MEG/company line management plan."),
        ("MOOR-005", "Mooring rope end-for-end / retirement records", ["end-for-end", "end for end", "retirement", "discard", "line management"], "Confirm end-for-end, retirement/discard and inspection records are updated and match HVPQ/company records."),
    ]
    hv_full = normalize_text(hvpq.get("_full_text", "")).lower()
    for cid, label, kws, action in mooring_checks:
        found = any(k.lower() in hv_full for k in kws)
        if found:
            out.append(Finding(cid, "Mooring", label, "INFO", "MEDIUM", hvpq_value=snippet(hvpq, "mooring_block"), reason=f"Related mooring keywords detected. Dates extracted near mooring block: {', '.join(fmt_date(x) for x in moor_dates[:5]) or 'None'}.", recommended_action=action, manual_verification_required="Yes"))
        else:
            out.append(Finding(cid, "Mooring", label, "MANUAL CHECK", "MEDIUM", hvpq_value="Keyword not clearly detected", reason="This is a recurring HVPQ observation area but reliable value extraction was not achieved.", recommended_action=action, manual_verification_required="Yes"))

    # 3. Piping pressure tests and pollution-prevention declarations.
    for cid, label, block_key, action in [
        ("PIPE-001", "Cargo/bunker piping pressure test declaration", "piping_pressure_block", "Confirm cargo and bunker piping hydrostatic pressure test interval, latest test date/pressure and supporting records."),
        ("POLL-001", "Overboard discharge / sea chest / scupper declaration", "pollution_block", "Confirm sea chest wording, overboard discharge blanks/testing arrangement and scupper/spill containment declarations against physical arrangement and records."),
        ("FIRE-001", "Foam / fixed firefighting declaration", "fire_foam_block", "Confirm foam type, last foam analysis date and fixed firefighting systems against certificates and onboard arrangement."),
        ("LIFT-001", "Lifting gear / crane declaration", "lifting_block", "Confirm crane/lifting gear SWL, annual/five-year test dates and certificates against HVPQ."),
        ("DIAG-001", "Mooring/manifold/fairlead/chock/bitt diagrams", "diagram_block", "Confirm diagrams are uploaded/available, current and consistent with HVPQ dimensions/equipment."),
    ]:
        if has_text(hvpq, block_key):
            out.append(Finding(cid, "Targeted HVPQ", label, "INFO", "MEDIUM", hvpq_value=snippet(hvpq, block_key), reason="Relevant HVPQ text detected; ship verification is still required because this is a recurring observation area.", recommended_action=action, manual_verification_required="Yes"))
        else:
            out.append(Finding(cid, "Targeted HVPQ", label, "MANUAL CHECK", "MEDIUM", hvpq_value="Not detected", reason="Relevant HVPQ text was not clearly detected. This may be extraction limitation or missing/blank declaration.", recommended_action=action, manual_verification_required="Yes"))

    # 4. Q88 cross-reference availability for commercial/vetting data.
    if q88:
        out.append(Finding("Q88-001", "Q88", "Q88 uploaded for cross-check", "INFO", "INFO", reason="Q88/reference file uploaded. Generic class/certificate comparison has been run where possible.", recommended_action="Use exported register to manually verify non-date equipment particulars from Q88 against HVPQ.", manual_verification_required="Yes"))
    else:
        out.append(Finding("Q88-001", "Q88", "Q88 not uploaded", "INFO", "INFO", reason="Q88 cross-check skipped because no Q88/reference file was uploaded.", recommended_action="Upload Q88 when available for vessel particulars/equipment cross-check."))

    return out

def rule_observation_library_scan(hvpq: Dict[str, Any], piq: Dict[str, Any], obs_docs: List[TextDoc]) -> List[Finding]:
    out: List[Finding] = []
    if not obs_docs:
        return out
    hv_text = " ".join([str(v) for v in hvpq.values() if isinstance(v, str)]).lower()
    piq_text = " ".join([str(v) for v in piq.values() if isinstance(v, str)] + piq.get("incident_blocks", [])).lower()

    # Use observation files as keyword/risk library, not as direct evidence.
    combined_obs = "\n".join(doc.full_text for doc in obs_docs)
    categories = classify_observation_patterns(combined_obs)
    for cat, data in categories.items():
        keywords = data["keywords"]
        # If the historical obs category is relevant but target docs don't contain expected keywords, flag targeted manual check.
        hv_hit = any(k.lower() in hv_text for k in keywords)
        piq_hit = any(k.lower() in piq_text for k in keywords)
        if data["priority"] >= 2 and not (hv_hit or piq_hit):
            out.append(Finding(
                f"OBS-{cat}", "Observation Library", data["label"], "MANUAL CHECK", data["risk"],
                reason="Historical observation pattern is high-risk, but matching target document text was not clearly detected.",
                recommended_action=data["action"],
                manual_verification_required="Yes",
            ))
    return out


def classify_observation_patterns(text: str) -> Dict[str, Dict[str, Any]]:
    tl = text.lower()
    patterns = {
        "INCIDENTS": {
            "label": "Incident non-reporting / mismatch pattern",
            "match_words": ["incident", "injury", "blackout", "failure", "reported", "not reported"],
            "keywords": ["incident", "injury", "failure", "blackout", "pollution", "grounding", "collision", "allision"],
            "risk": "HIGH",
            "priority": 3,
            "action": "Ship to confirm PIQ/HVPQ incident declarations are complete; if no incidents are entered, provide positive confirmation.",
        },
        "CII": {
            "label": "CII / EEXI / EEDI accuracy pattern",
            "match_words": ["cii", "eexi", "eedi", "aer"],
            "keywords": ["cii", "eexi", "eedi", "aer"],
            "risk": "HIGH",
            "priority": 3,
            "action": "Verify HVPQ environmental values against latest CII/EEXI/EEDI evidence.",
        },
        "CLASS_CERT": {
            "label": "Class / certificate date accuracy pattern",
            "match_words": ["certificate", "class", "dry dock", "iws", "survey", "expiry", "issued"],
            "keywords": ["certificate", "class", "dry dock", "iws", "survey", "expiry"],
            "risk": "HIGH",
            "priority": 3,
            "action": "Verify HVPQ certificate and survey dates against latest certificates/Class Status.",
        },
        "PSC": {
            "label": "PSC history accuracy pattern",
            "match_words": ["psc", "port state", "deficien", "detention"],
            "keywords": ["psc", "port state", "detention", "deficien"],
            "risk": "HIGH",
            "priority": 3,
            "action": "Verify last three PSC entries, port, MOU, deficiencies, detention and OCIMF PSC database entry.",
        },
        "MOORING": {
            "label": "Mooring / BHC / brake test pattern",
            "match_words": ["mooring", "brake", "bhc", "split drum", "rope"],
            "keywords": ["mooring", "brake", "bhc", "split drum", "rope"],
            "risk": "MEDIUM",
            "priority": 2,
            "action": "Verify mooring winch, BHC, brake test, line data and diagrams against records.",
        },
        "TANKS": {
            "label": "Tank coating / inspection pattern",
            "match_words": ["tank", "coating", "ballast", "cargo tank", "void"],
            "keywords": ["tank", "coating", "ballast", "cargo tank", "void"],
            "risk": "MEDIUM",
            "priority": 2,
            "action": "Verify cargo/ballast/void inspection dates, coating condition and tank list completeness.",
        },
    }
    found = {}
    for key, cfg in patterns.items():
        count = sum(tl.count(w) for w in cfg["match_words"])
        if count > 0:
            found[key] = cfg
    return found


# -----------------------------
# Export helpers
# -----------------------------

def findings_to_df(findings: List[Finding]) -> pd.DataFrame:
    rows = []
    for f in findings:
        d = asdict(f)
        d["score"] = f.score
        rows.append(d)
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    cols = [
        "check_id", "module", "check", "status", "risk", "score", "hvpq_value", "piq_value",
        "class_status_value", "q88_value", "reason", "recommended_action", "manual_verification_required", "source_location",
    ]
    return df[cols]


def make_excel_download(df: pd.DataFrame, summary: Dict[str, Any]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Audit Register", index=False)
        workbook = writer.book
        ws = writer.sheets["Audit Register"]
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True})
        wrap_fmt = workbook.add_format({"text_wrap": True, "valign": "top"})
        high_fmt = workbook.add_format({"bg_color": "#F8CBAD"})
        crit_fmt = workbook.add_format({"bg_color": "#FF9999", "bold": True})
        pass_fmt = workbook.add_format({"bg_color": "#C6E0B4"})
        for col, name in enumerate(df.columns):
            ws.write(0, col, name, header_fmt)
            width = 18
            if name in ["reason", "recommended_action", "source_location"]:
                width = 45
            if name in ["hvpq_value", "piq_value", "class_status_value", "q88_value"]:
                width = 32
            ws.set_column(col, col, width, wrap_fmt)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
        risk_col = df.columns.get_loc("risk")
        ws.conditional_format(1, risk_col, len(df), risk_col, {"type": "text", "criteria": "containing", "value": "CRITICAL", "format": crit_fmt})
        ws.conditional_format(1, risk_col, len(df), risk_col, {"type": "text", "criteria": "containing", "value": "HIGH", "format": high_fmt})
        ws.conditional_format(1, risk_col, len(df), risk_col, {"type": "text", "criteria": "containing", "value": "PASS", "format": pass_fmt})

        summary_df = pd.DataFrame(list(summary.items()), columns=["Metric", "Value"])
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        ws2 = writer.sheets["Summary"]
        ws2.set_column(0, 0, 35)
        ws2.set_column(1, 1, 25)
        for col, name in enumerate(summary_df.columns):
            ws2.write(0, col, name, header_fmt)
    return output.getvalue()


def summarize_findings(df: pd.DataFrame) -> Dict[str, Any]:
    if df.empty:
        return {"Total checks": 0}
    return {
        "Total checks": len(df),
        "Critical": int((df["risk"] == "CRITICAL").sum()),
        "High": int((df["risk"] == "HIGH").sum()),
        "Medium": int((df["risk"] == "MEDIUM").sum()),
        "Manual checks": int((df["manual_verification_required"] == "Yes").sum()),
        "Pass": int((df["risk"] == "PASS").sum()),
        "Overall risk score": int(df["score"].sum()),
    }


# -----------------------------
# Streamlit UI
# -----------------------------

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption("Observation-driven, rule-based checks. Output is a targeted ship-verification register, not an automatic certification.")

    with st.sidebar:
        st.header("Upload documents")
        hvpq_file = st.file_uploader("HVPQ PDF / Q88-style vessel particulars", type=["pdf", "xlsx", "xls", "csv", "docx"], key="hvpq")
        piq_file = st.file_uploader("PIQ PDF", type=["pdf", "xlsx", "xls", "csv", "docx"], key="piq")
        class_file = st.file_uploader("Class Status / Class survey status / certificate status", type=["pdf", "xlsx", "xls", "csv", "docx"], key="class_status")
        q88_file = st.file_uploader("Optional Q88 / other VPQ reference", type=["pdf", "xlsx", "xls", "csv", "docx"], key="q88")
        obs_files = st.file_uploader("Optional observation libraries", type=["xlsx", "xls", "csv", "pdf", "docx"], accept_multiple_files=True, key="obs")

        st.header("Rules")
        ts_gap = st.number_input("Technical Superintendent max gap months", value=DEFAULT_TS_GAP_MONTHS, min_value=1.0, step=0.5)
        ms_gap = st.number_input("Marine Superintendent max gap months", value=DEFAULT_MS_GAP_MONTHS, min_value=1.0, step=0.5)
        cert_warning = st.number_input("Certificate expiry warning days", value=DEFAULT_CERT_WARNING_DAYS, min_value=1, step=15)
        survey_warning = st.number_input("Survey/drydock due warning days", value=DEFAULT_SURVEY_WARNING_DAYS, min_value=1, step=30)
        run_btn = st.button("Run checks", type="primary")

    if not run_btn:
        st.info("Upload HVPQ and PIQ at minimum, then click **Run checks**. Class Status and Q88 are optional but recommended.")
        st.markdown(
            """
            **Core logic included:**
            - HVPQ vs PIQ mismatch checks
            - Incident declaration checks only from HVPQ/PIQ; no JiBe incident list required
            - If no incidents are entered, the app flags a positive confirmation requirement
            - Class Status cross-check using generic extraction because class society formats differ
            - Optional Q88 reference comparison
            - Exportable ship-verification register
            """
        )
        return

    if not hvpq_file or not piq_file:
        st.error("Please upload at least HVPQ and PIQ.")
        return

    with st.spinner("Reading files and running rules..."):
        hvpq_doc = read_uploaded_file(hvpq_file.name, hvpq_file.getvalue())
        piq_doc = read_uploaded_file(piq_file.name, piq_file.getvalue())
        class_doc = read_uploaded_file(class_file.name, class_file.getvalue()) if class_file else None
        q88_doc = read_uploaded_file(q88_file.name, q88_file.getvalue()) if q88_file else None
        obs_docs = [read_uploaded_file(f.name, f.getvalue()) for f in obs_files] if obs_files else []

        hvpq = extract_hvpq_fields(hvpq_doc)
        piq = extract_piq_fields(piq_doc)
        class_status = extract_reference_fields(class_doc, "Class Status") if class_doc else {}
        q88 = extract_reference_fields(q88_doc, "Q88") if q88_doc else {}
        settings = {
            "ts_gap_months": ts_gap,
            "ms_gap_months": ms_gap,
            "cert_warning_days": cert_warning,
            "survey_warning_days": survey_warning,
        }
        findings = run_all_rules(hvpq, piq, class_status, q88, settings, obs_docs)
        df = findings_to_df(findings)
        summary = summarize_findings(df)

    st.subheader("Dashboard")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total", summary.get("Total checks", 0))
    c2.metric("Critical", summary.get("Critical", 0))
    c3.metric("High", summary.get("High", 0))
    c4.metric("Medium", summary.get("Medium", 0))
    c5.metric("Manual", summary.get("Manual checks", 0))
    c6.metric("Score", summary.get("Overall risk score", 0))

    st.subheader("Extracted identity")
    id_df = pd.DataFrame([
        {"Document": "HVPQ", "Name": hvpq.get("vessel_name", ""), "IMO": hvpq.get("imo", ""), "Date": fmt_date(hvpq.get("doc_date"))},
        {"Document": "PIQ", "Name": piq.get("vessel_name", ""), "IMO": "", "Date": fmt_date(piq.get("doc_date"))},
    ])
    st.dataframe(id_df, use_container_width=True, hide_index=True)

    st.subheader("Audit register")
    risk_filter = st.multiselect("Filter risk", sorted(df["risk"].unique().tolist()) if not df.empty else [], default=sorted(df["risk"].unique().tolist()) if not df.empty else [])
    view_df = df[df["risk"].isin(risk_filter)] if risk_filter and not df.empty else df
    st.dataframe(view_df, use_container_width=True, hide_index=True)

    excel_bytes = make_excel_download(df, summary)
    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button("Download Excel audit register", excel_bytes, file_name="hvpq_piq_audit_register.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download CSV audit register", csv_bytes, file_name="hvpq_piq_audit_register.csv", mime="text/csv")

    with st.expander("Raw extracted fields - HVPQ"):
        st.json({k: (fmt_date(v) if isinstance(v, date) else str(v)[:500]) for k, v in hvpq.items() if k != "certificates"})
    with st.expander("Raw extracted fields - PIQ"):
        st.json({k: (", ".join(fmt_date(d) for d in v) if isinstance(v, list) and all(isinstance(x, date) for x in v) else str(v)[:500]) for k, v in piq.items()})
    with st.expander("Important limitation"):
        st.write(
            "Class status documents vary by classification society and layout. The app therefore uses generic date/keyword proximity extraction and deliberately marks uncertain items as manual checks. For high accuracy, future versions should add class-specific parsers for NK, LR, DNV, ABS, BV, KR, RINA, CCS, etc."
        )


if __name__ == "__main__":
    main()
