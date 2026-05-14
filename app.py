from __future__ import annotations

import io
import re
import json
import zipfile
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from datetime import datetime, date
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import streamlit as st
from dateutil import parser as dateparser
from dateutil.relativedelta import relativedelta

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

# -----------------------------
# Data structures
# -----------------------------

@dataclass
class DocPack:
    name: str
    kind: str
    text: str = ""
    fields: Dict[str, Any] = None
    qmap: Dict[str, str] = None
    meta: Dict[str, Any] = None

    def __post_init__(self):
        self.fields = self.fields or {}
        self.qmap = self.qmap or {}
        self.meta = self.meta or {}

@dataclass
class Finding:
    area: str
    check: str
    status: str
    risk: str
    hvpq_value: str = ""
    piq_value: str = ""
    class_value: str = ""
    q88_value: str = ""
    xml_value: str = ""
    reason: str = ""
    required_action: str = ""
    source: str = ""
    question_ref: str = ""

# -----------------------------
# Config / Targeted rule library
# -----------------------------

RISK_ORDER = {"CRITICAL": 5, "HIGH": 4, "MEDIUM": 3, "LOW": 2, "INFO": 1, "PASS": 0}
ACTIONABLE = {"CRITICAL", "HIGH", "MEDIUM", "MANUAL CHECK", "WARNING"}

# Field specs use aliases because PIQ/HVPQ/Class/Q88 labels will never be identical.
# Keep them objective; do not overinterpret.
FIELD_SPECS: Dict[str, Dict[str, Any]] = {
    "vessel_name": {
        "area": "Identity", "risk": "HIGH", "kind": "text", "loose": True,
        "aliases": ["Name of ship", "Vessel Name", "Ship Name", "Vessel", "Name"]
    },
    "imo_number": {
        "area": "Identity", "risk": "CRITICAL", "kind": "number",
        "aliases": ["LR/IMO number", "IMO/LR Number", "IMO number", "IMO No", "IMO"]
    },
    "document_date": {
        "area": "Document status", "risk": "MEDIUM", "kind": "date",
        "aliases": ["Date this HVPQ document completed", "Date", "Document exported", "completed"]
    },
    "vessel_type": {
        "area": "General Particulars", "risk": "MEDIUM", "kind": "semantic_text", "loose": True,
        "aliases": ["Vessel Type", "type of ship", "type of tanker", "ship type", "Product Carrier", "Products/Chemical Tanker"]
    },
    "flag": {"area": "General Particulars", "risk": "MEDIUM", "kind": "text", "aliases": ["Flag", "Flag State"]},
    "class_society": {"area": "Class", "risk": "HIGH", "kind": "semantic_text", "aliases": ["Classification Society", "Class Society", "Class", "Recognized Organization"]},
    "last_drydock": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date of last dry dock", "last dry dock", "last drydock", "dry dock date", "docking survey"]},
    "next_drydock_due": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date next dry dock due", "next dry dock due", "next drydock", "docking survey due"]},
    "last_iws": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date of last IWS", "last IWS", "in water survey", "underwater survey"]},
    "next_iws_due": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date next IWS due", "next IWS due", "IWS due", "underwater survey due"]},
    "last_special_survey": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date of last special survey", "last special survey", "special survey date"]},
    "next_special_survey_due": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date next special survey due", "next special survey due", "special survey due"]},
    "last_annual_survey": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date of last annual survey", "last annual survey", "annual survey"]},
    "last_intermediate_survey": {"area": "Class / Survey", "risk": "HIGH", "kind": "date", "aliases": ["Date of Last Intermediate survey", "last intermediate survey", "intermediate survey"]},
    "conditions_of_class": {"area": "Class / Survey", "risk": "CRITICAL", "kind": "yesno", "aliases": ["open Conditions of Class", "Conditions of Class", "COC", "Condition of Class"]},
    "memoranda_of_class": {"area": "Class / Survey", "risk": "HIGH", "kind": "yesno", "aliases": ["Memoranda of Class", "Memorandum of Class", "MOC"]},
    "flag_dispensation": {"area": "Class / Survey", "risk": "HIGH", "kind": "yesno", "aliases": ["flag state dispensations", "dispensation", "exemption"]},
    "cii_rating": {"area": "Environmental", "risk": "HIGH", "kind": "semantic_text", "aliases": ["CII rating", "Carbon Intensity Indicator", "attained annual operational CII", "CII"]},
    "cii_verified_by": {"area": "Environmental", "risk": "MEDIUM", "kind": "semantic_text", "aliases": ["CII rating verified", "CII verified", "Is the CII rating verified", "verified by"]},
    "eexi_rating": {"area": "Environmental", "risk": "HIGH", "kind": "number", "aliases": ["EEXI rating", "Energy Efficiency Existing ship Index", "EEXI"]},
    "eedi_rating": {"area": "Environmental", "risk": "MEDIUM", "kind": "number", "aliases": ["EEDI rating", "Energy Efficiency Design Index", "EEDI"]},
    "eiv_rating": {"area": "Environmental", "risk": "MEDIUM", "kind": "number", "aliases": ["EIV rating", "Estimated Index Value", "EIV"]},
    "last_psc_date": {"area": "PSC", "risk": "HIGH", "kind": "date", "aliases": ["Date of last Port State Control Inspection", "Date of PSC Inspection", "last PSC", "PSC inspection"]},
    "last_psc_port": {"area": "PSC", "risk": "HIGH", "kind": "semantic_text", "aliases": ["Port of last Port State Control Inspection", "port did the inspection take place", "last PSC port", "Port"]},
    "psc_detained": {"area": "PSC", "risk": "CRITICAL", "kind": "yesno", "aliases": ["detained during the last 36 months", "Was the vessel detained", "detained"]},
    "incident_pollution": {"area": "Incidents", "risk": "HIGH", "kind": "yesno", "aliases": ["pollution incident", "release to the environment", "MARPOL"]},
    "incident_grounding": {"area": "Incidents", "risk": "CRITICAL", "kind": "yesno", "aliases": ["grounding", "hard aground", "touched bottom", "suspected of touching bottom"]},
    "incident_other": {"area": "Incidents", "risk": "HIGH", "kind": "yesno", "aliases": ["other incidents during the past 12 months", "any other incidents", "incident during previous 12 months"]},
    "foam_type": {"area": "Firefighting", "risk": "MEDIUM", "kind": "semantic_text", "aliases": ["type of foam", "fixed foam", "foam type", "AR-AFFF"]},
    "foam_test_date": {"area": "Firefighting", "risk": "HIGH", "kind": "date", "aliases": ["date of supply of the foam", "last Test Analysis Certificate", "foam test", "foam analysis"]},
    "machinery_space_fixed_fire": {"area": "Firefighting", "risk": "MEDIUM", "kind": "semantic_text", "aliases": ["machinery space fixed", "engine room", "fixed fire extinguishing system - machinery space"]},
    "cargo_pumproom_fixed_fire": {"area": "Firefighting", "risk": "MEDIUM", "kind": "semantic_text", "aliases": ["cargo pumproom", "Cargo pump room", "fixed fire extinguishing system - cargo pumproom"]},
    "cargo_pressure_test": {"area": "Pollution / Cargo", "risk": "HIGH", "kind": "yesno", "aliases": ["cargo piping pressure tests", "hydrostatically pressure test cargo piping", "cargo pressure test"]},
    "cargo_pressure": {"area": "Pollution / Cargo", "risk": "MEDIUM", "kind": "number", "aliases": ["cargo piping pressure", "specify pressure", "cargo pressure"]},
    "bunker_pressure_test": {"area": "Pollution / Cargo", "risk": "HIGH", "kind": "yesno", "aliases": ["bunker piping pressure tests", "hydrostatically pressure test bunker piping", "bunker pressure test"]},
    "bunker_pressure": {"area": "Pollution / Cargo", "risk": "MEDIUM", "kind": "number", "aliases": ["bunker piping pressure", "bunker pressure"]},
    "overboard_blanks": {"area": "Pollution / Cargo", "risk": "HIGH", "kind": "yesno", "aliases": ["overboard discharges fitted with blanks", "overboard valves", "testing arrangement for the overboard valves"]},
    "sea_chest": {"area": "Pollution / Cargo", "risk": "MEDIUM", "kind": "semantic_text", "aliases": ["sea valves", "seachest", "sea chest", "cargo sea chest"]},
    "cargo_tank_oldest_inspection": {"area": "Tank Inspection", "risk": "HIGH", "kind": "date", "aliases": ["oldest inspection report for all cargo and slop", "cargo tank inspection", "cargo and slop tanks"]},
    "ballast_tank_oldest_inspection": {"area": "Tank Inspection", "risk": "HIGH", "kind": "date", "aliases": ["oldest inspection report for all ballast", "ballast tank inspection", "ballast tanks"]},
    "void_space_oldest_inspection": {"area": "Tank Inspection", "risk": "HIGH", "kind": "date", "aliases": ["oldest inspection report for all void", "void space inspection", "void spaces"]},
    "tank_inspection_frequency": {"area": "Tank Inspection", "risk": "MEDIUM", "kind": "number", "aliases": ["required frequency of inspection for cargo tanks", "required frequency", "Frequency of Inspections"]},
    "msmp": {"area": "Mooring", "risk": "HIGH", "kind": "yesno", "aliases": ["Mooring System Management Plan", "MSMP"]},
    "lmp": {"area": "Mooring", "risk": "HIGH", "kind": "yesno", "aliases": ["Line Management Plan", "LMP"]},
    "sdmbl": {"area": "Mooring", "risk": "MEDIUM", "kind": "number", "aliases": ["Ship design MBL", "SDMBL", "design MBL"]},
    "brake_testing_equipment": {"area": "Mooring", "risk": "HIGH", "kind": "yesno", "aliases": ["brake testing equipment", "brake testing equipment on board"]},
    "brake_test_date": {"area": "Mooring", "risk": "HIGH", "kind": "date", "aliases": ["last brake", "brake holding capacity test", "BHC test", "brake test"]},
    "split_drum": {"area": "Mooring", "risk": "MEDIUM", "kind": "yesno", "aliases": ["split drum", "split drums"]},
    "rope_end_for_end": {"area": "Mooring", "risk": "MEDIUM", "kind": "date", "aliases": ["end-for-end", "end for end", "rope rotation", "mooring rope end"]},
    "mooring_ropes": {"area": "Mooring", "risk": "MEDIUM", "kind": "semantic_text", "aliases": ["Mooring ropes", "mooring lines", "tails", "Line Management"]},
    "manifold_diagram": {"area": "Diagrams", "risk": "MEDIUM", "kind": "presence", "aliases": ["Manifold arrangement diagram", "manifold diagram", "manifold arrangement"]},
    "mooring_diagram": {"area": "Diagrams", "risk": "MEDIUM", "kind": "presence", "aliases": ["mooring winch layout", "mooring arrangement diagram", "fairlead", "chock", "bitt diagram"]},
    "crane_test_date": {"area": "Lifting Gear", "risk": "HIGH", "kind": "date", "aliases": ["last 5yr test", "crane", "lifting gear", "annual thorough examination", "five yearly test"]},
    "generator_count": {"area": "Machinery", "risk": "MEDIUM", "kind": "number", "aliases": ["power generators", "generator details", "number of generators", "auxiliary generators"]},
}

# Targeted objective checks based on common observation themes. These are ship-facing.
TARGETED_CHECKS = [
    ("Mooring", "Brake test date / BHC / rendering load", ["brake", "brake holding", "BHC", "rendering load"], "Confirm latest brake test certificate/date, BHC and rendering load match HVPQ/Q88/records."),
    ("Mooring", "Mooring ropes / tails / end-for-end / retirement", ["rope", "tail", "end-for-end", "end for end", "retirement", "discard", "LMP"], "Confirm ropes/tails particulars, certificate dates, end-for-end and retirement dates match LMP and HVPQ."),
    ("Tank Inspection", "Cargo / ballast / void inspection sequence", ["cargo tank", "ballast", "void", "coating", "inspection"], "Confirm oldest inspection dates and annual frequency against inspection records."),
    ("Class / Certificates", "Class status and survey dates", ["class", "dry dock", "IWS", "special survey", "annual survey", "intermediate"], "Verify HVPQ survey dates and COC/MOC/dispensation against latest Class Status."),
    ("Certificates", "Certificate issue/expiry/endorsement dates", ["certificate", "expiry", "expires", "endorsement", "SMC", "ISSC", "IOPP", "Loadline", "Safety Equipment"], "Verify certificate dates against certificate copies/Class Status/Q88."),
    ("PSC", "Last PSC date/port/MOU/deficiencies/detention", ["PSC", "Port State", "deficiencies", "detained", "MOU"], "Verify last three PSC entries match PIQ, HVPQ and company PSC records."),
    ("Incidents", "Incident declaration status", ["incident", "injury", "blackout", "failure", "pollution", "grounding", "mooring rope parting"], "Confirm whether any incident is declared in HVPQ/PIQ; if none, ship must positively confirm nil incidents."),
    ("Pollution / Cargo", "Overboard blanks / sea chest / pressure tests", ["overboard", "sea chest", "scupper", "cargo piping", "bunker piping", "pressure test"], "Verify arrangement and latest pressure test records match HVPQ."),
    ("Firefighting", "Foam / fixed systems / rescue boat", ["foam", "fixed firefighting", "rescue boat", "lifeboat", "sample locker"], "Verify fixed systems and foam test date against certificates/onboard arrangements."),
    ("Cargo / IGS", "IGS / venting / heating / cargo pumps", ["IGS", "inert", "vent", "heating", "cargo pump", "vapour", "COW"], "Verify cargo system declarations against manuals/diagrams/vessel arrangement."),
    ("Diagrams", "Mooring / manifold / fairlead / chock / bitt diagrams", ["diagram", "manifold", "fairlead", "chock", "bitt", "SPM"], "Confirm diagrams are attached/current and match physical arrangement."),
    ("Lifting Gear", "Crane annual / five-year test", ["crane", "lifting", "SWL", "5yr", "five-year", "annual"], "Verify crane/lifting gear certificate dates and SWL match HVPQ/Q88."),
]

# Known semantic equivalences that should not be treated as mismatches.
EQUIVALENCE_GROUPS = [
    {"product carrier", "products tanker", "product tanker", "products/chemical tanker", "oil/product tanker", "chemical/products tanker"},
    {"no", "not applicable", "na", "n/a", "not fitted", "none"},
    {"yes", "fitted", "provided", "available"},
    {"nippon kaiji kyokai", "classnk", "class nk", "nk"},
    {"co2", "carbon dioxide"},
]

# -----------------------------
# Text / extraction utilities
# -----------------------------

def normalize_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\x00", " ").replace("\u00a0", " ")
    s = re.sub(r"[\t\r]+", " ", s)
    s = re.sub(r" +", " ", s)
    return s.strip()


def clean_for_compare(s: Any) -> str:
    s = normalize_text(s).lower()
    s = re.sub(r"[^a-z0-9./+\- ]+", " ", s)
    s = re.sub(r"\b(lr2|mr|aframax|suezmax)\b", " ", s)  # extraction noise / class descriptions not material for vessel name
    s = re.sub(r"\s+", " ", s).strip()
    return s


def token_sort(s: Any) -> str:
    return " ".join(sorted(clean_for_compare(s).split()))


def fuzzy_ratio(a: Any, b: Any) -> float:
    aa, bb = clean_for_compare(a), clean_for_compare(b)
    if not aa or not bb:
        return 0.0
    return SequenceMatcher(None, aa, bb).ratio()


def similar_enough(a: Any, b: Any, kind: str = "text", loose: bool = False) -> bool:
    aa, bb = clean_for_compare(a), clean_for_compare(b)
    if not aa or not bb:
        return False
    if aa == bb:
        return True
    # Equivalence groups
    for group in EQUIVALENCE_GROUPS:
        if any(g in aa for g in group) and any(g in bb for g in group):
            return True
    if kind == "semantic_text" or loose:
        # tolerate product carrier vs products/chemical tanker and naming suffixes
        if aa in bb or bb in aa:
            return True
        return max(fuzzy_ratio(aa, bb), fuzzy_ratio(token_sort(aa), token_sort(bb))) >= (0.72 if loose else 0.82)
    return max(fuzzy_ratio(aa, bb), fuzzy_ratio(token_sort(aa), token_sort(bb))) >= (0.90 if not loose else 0.75)


def parse_date(value: Any) -> Optional[date]:
    s = normalize_text(value)
    if not s:
        return None
    # protect question numbers / decimals from being parsed as dates
    if re.fullmatch(r"\d+(?:\.\d+){1,4}", s):
        return None
    if re.fullmatch(r"\d+(?:\.\d+)?", s):
        return None
    s = re.sub(r"\b(UTC|LT|hrs?|hours?)\b", " ", s, flags=re.I)
    # ISO inside XML/PDF
    m = re.search(r"\b(20\d{2}|19\d{2})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            pass
    # Common OCR text dates
    m = re.search(r"\b(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(20\d{2}|19\d{2})\b", s, flags=re.I)
    if m:
        try:
            return dateparser.parse(m.group(0), dayfirst=True).date()
        except Exception:
            return None
    m = re.search(r"\b(\d{1,2})[-/](\d{1,2})[-/](20\d{2}|19\d{2})\b", s)
    if m:
        try:
            return dateparser.parse(m.group(0), dayfirst=True).date()
        except Exception:
            return None
    try:
        dt = dateparser.parse(s, fuzzy=True, dayfirst=True)
        if dt and 1990 <= dt.year <= 2050:
            return dt.date()
    except Exception:
        return None
    return None


def parse_number(value: Any) -> Optional[float]:
    s = normalize_text(value)
    if not s:
        return None
    # avoid question IDs
    if re.fullmatch(r"\d+(?:\.\d+){2,5}", s):
        return None
    m = re.search(r"[-+]?\d{1,3}(?:,\d{3})*(?:\.\d+)?|[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except Exception:
        return None


def parse_yesno(value: Any) -> Optional[bool]:
    s = clean_for_compare(value)
    if not s:
        return None
    # avoid matching 'not detained? no' incorrectly is okay as value only ideally.
    if re.search(r"\b(yes|true|1)\b", s):
        return True
    if re.search(r"\b(no|false|0|not applicable|n/a|na|none|not fitted)\b", s):
        return False
    return None


def extract_pdf_text(file_bytes: bytes, filename: str = "") -> str:
    text = ""
    if fitz is not None:
        try:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            pages = []
            for i, p in enumerate(doc):
                pages.append(f"\n--- PAGE {i+1} ---\n" + p.get_text("text"))
            text = "\n".join(pages)
        except Exception:
            text = ""
    if not text and PdfReader is not None:
        try:
            reader = PdfReader(io.BytesIO(file_bytes))
            text = "\n".join(f"\n--- PAGE {i+1} ---\n" + (p.extract_text() or "") for i, p in enumerate(reader.pages))
        except Exception:
            text = ""
    return normalize_text_keep_lines(text)


def normalize_text_keep_lines(s: str) -> str:
    s = s.replace("\r", "\n").replace("\u00a0", " ")
    s = unicodedata.normalize("NFKC", s)
    lines = []
    for line in s.splitlines():
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def read_any_upload(uploaded, kind: str) -> DocPack:
    if uploaded is None:
        return DocPack(name="", kind=kind)
    b = uploaded.getvalue()
    name = uploaded.name
    lower = name.lower()
    if lower.endswith(".pdf"):
        return DocPack(name=name, kind=kind, text=extract_pdf_text(b, name))
    if lower.endswith(".xml"):
        fields, meta = parse_hvpq_xml(b)
        return DocPack(name=name, kind=kind, text=xml_to_searchable_text(fields, meta), fields=fields, meta=meta)
    if lower.endswith(".txt"):
        return DocPack(name=name, kind=kind, text=normalize_text_keep_lines(b.decode("utf-8", errors="ignore")))
    if lower.endswith((".xlsx", ".xls", ".csv")):
        text = read_tabular_as_text(b, lower)
        return DocPack(name=name, kind=kind, text=text)
    return DocPack(name=name, kind=kind, text="")


def read_tabular_as_text(b: bytes, lower: str) -> str:
    try:
        if lower.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(b), dtype=str)
        else:
            df = pd.read_excel(io.BytesIO(b), dtype=str)
        df = df.fillna("")
        rows = [" | ".join(map(str, df.columns.tolist()))]
        for _, row in df.head(2000).iterrows():
            rows.append(" | ".join(normalize_text(x) for x in row.tolist()))
        return "\n".join(rows)
    except Exception:
        return ""


def parse_hvpq_xml(b: bytes, ctrl_map: Optional[pd.DataFrame] = None) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """Parse OCIMF XPQ response XML. Without control mapping this is intentionally limited.
    The response XML normally contains ctrl GUIDs and response values, but not the human readable question labels.
    If a ctrl->qid/label map is supplied, the responses become directly usable.
    """
    fields: Dict[str, Any] = {}
    meta: Dict[str, Any] = {"xml_warning": "Response XML parsed. Control IDs need template mapping for question-level checks."}
    try:
        root = ET.fromstring(b)
    except Exception as e:
        return {}, {"error": f"XML parse failed: {e}"}
    ns = ""
    if root.tag.startswith("{"):
        ns = root.tag.split("}")[0].strip("{")
    def q(name): return f"{{{ns}}}{name}" if ns else name
    header = root.find(q("Header"))
    if header is not None:
        vessel = header.find(q("Vessel"))
        templ = header.find(q("Template"))
        doc = header.find(q("Document"))
        if vessel is not None:
            meta["vessel_name"] = vessel.attrib.get("name", "")
            meta["imo_number"] = vessel.attrib.get("id", "")
            fields["vessel_name"] = meta["vessel_name"]
            fields["imo_number"] = meta["imo_number"]
        if templ is not None:
            meta["template_code"] = templ.attrib.get("code", "")
            meta["template_variant"] = templ.attrib.get("variant", "")
        if doc is not None:
            meta["document_name"] = doc.attrib.get("name", "")
            meta["exported"] = doc.attrib.get("exported", "")
            fields["document_date"] = meta["exported"]
    responses = []
    for resp in root.findall(q("Response")):
        ctrl = resp.attrib.get("ctrl", "")
        typ = resp.attrib.get("type", "")
        val = response_value(resp, ns)
        responses.append({"ctrl": ctrl, "type": typ, "value": val})
    meta["response_count"] = len(responses)
    meta["response_types"] = dict(pd.Series([r["type"] for r in responses]).value_counts()) if responses else {}
    fields["_raw_responses"] = responses
    if ctrl_map is not None and not ctrl_map.empty:
        mapped = apply_ctrl_map(responses, ctrl_map)
        fields.update(mapped)
        meta["xml_warning"] = "Response XML parsed with supplied ctrl mapping."
    return fields, meta


def response_value(elem: ET.Element, ns: str) -> str:
    def q(name): return f"{{{ns}}}{name}" if ns else name
    for tag in ["ResponseString", "ResponseDate", "ResponseBoolean", "ResponseInt", "ResponseDecimal", "ResponseMemo"]:
        x = elem.find(q(tag))
        if x is not None and x.text is not None:
            if tag == "ResponseBoolean":
                return "Yes" if x.text.strip() in {"1", "true", "True"} else "No"
            return normalize_text(x.text)
    selected = elem.find(q("SelectedItems"))
    if selected is not None:
        vals = []
        for item in selected.findall(q("SelectedItem")):
            rs = item.find(q("ResponseString"))
            if rs is not None and rs.text:
                vals.append(normalize_text(rs.text))
        if vals:
            return ", ".join(vals)
    # CellResponse tables
    cells = []
    for c in elem.iter():
        if c.tag.split("}")[-1] == "CellResponse":
            val = response_value(c, ns)
            if val:
                cells.append(val)
    return " | ".join(cells[:50])


def apply_ctrl_map(responses: List[Dict[str, str]], ctrl_map: pd.DataFrame) -> Dict[str, Any]:
    cols = {clean_for_compare(c): c for c in ctrl_map.columns}
    ctrl_col = cols.get("ctrl") or cols.get("control") or cols.get("control id") or cols.get("controlid")
    qid_col = cols.get("qid") or cols.get("question id") or cols.get("question") or cols.get("ref")
    label_col = cols.get("label") or cols.get("question label") or cols.get("text") or cols.get("question text")
    if not ctrl_col:
        return {}
    map_by_ctrl = {}
    for _, r in ctrl_map.iterrows():
        ctrl = normalize_text(r.get(ctrl_col, ""))
        if not ctrl:
            continue
        qid = normalize_text(r.get(qid_col, "")) if qid_col else ""
        label = normalize_text(r.get(label_col, "")) if label_col else ""
        key = qid or label or ctrl
        map_by_ctrl[ctrl.upper()] = key
    out = {}
    for r in responses:
        key = map_by_ctrl.get(r["ctrl"].upper())
        if key:
            out[key] = r["value"]
    return out


def xml_to_searchable_text(fields: Dict[str, Any], meta: Dict[str, Any]) -> str:
    lines = []
    for k, v in meta.items():
        lines.append(f"{k}: {v}")
    for k, v in fields.items():
        if k == "_raw_responses":
            continue
        lines.append(f"{k}: {v}")
    return "\n".join(lines)

# -----------------------------
# HVPQ/PIQ parsing helpers
# -----------------------------

def build_question_blocks(text: str) -> Dict[str, str]:
    """Build approximate blocks keyed by HVPQ/PIQ question number from PDF text."""
    blocks: Dict[str, List[str]] = {}
    current = None
    for line in text.splitlines():
        # HVPQ: 1.1.1 Date... ; PIQ sometimes: 2.2.1001.
        m = re.match(r"^((?:\d+\.){1,4}\d+)\.?\s+(.*)", line)
        if m:
            qid = m.group(1).rstrip(".")
            current = qid
            blocks.setdefault(qid, []).append(m.group(2).strip())
        elif current:
            blocks[current].append(line.strip())
    return {k: "\n".join(v).strip() for k, v in blocks.items()}


def extract_near_alias(text: str, aliases: Iterable[str], kind: str = "text", window_lines: int = 4) -> str:
    if not text:
        return ""
    lines = text.splitlines()
    best = (0.0, "")
    for i, line in enumerate(lines):
        cl = clean_for_compare(line)
        for alias in aliases:
            ca = clean_for_compare(alias)
            if not ca:
                continue
            score = 0.0
            if ca in cl:
                score = 1.0
            else:
                score = fuzzy_ratio(ca, cl[:max(len(ca)+40, 60)])
            if score >= 0.72:
                chunk = "\n".join(lines[i:i+window_lines])
                value = strip_label_value(chunk, alias, kind)
                if value and score > best[0]:
                    best = (score, value)
    return best[1]


def strip_label_value(chunk: str, alias: str, kind: str) -> str:
    # try same-line extraction first
    lines = chunk.splitlines()
    for idx, line in enumerate(lines):
        if clean_for_compare(alias) in clean_for_compare(line) or fuzzy_ratio(alias, line) > 0.72:
            # remove leading question number and alias-like text
            candidate = line
            candidate = re.sub(r"^\d+(?:\.\d+)*\.?\s*", "", candidate)
            # If alias appears literally, keep after it
            m = re.search(re.escape(alias), candidate, flags=re.I)
            if m:
                candidate = candidate[m.end():]
            else:
                # split on ? if present
                if "?" in candidate:
                    candidate = candidate.split("?")[-1]
            candidate = candidate.strip(" :-?\t")
            if usable_value(candidate, kind):
                return candidate
            # else next line(s)
            for nxt in lines[idx+1:idx+4]:
                nxt2 = nxt.strip(" :-?\t")
                if usable_value(nxt2, kind):
                    return nxt2
    # fallback by kind in whole chunk
    if kind == "date":
        d = parse_date(chunk)
        return d.isoformat() if d else ""
    if kind == "number":
        n = parse_number(chunk)
        return str(n) if n is not None else ""
    if kind == "yesno":
        y = parse_yesno(chunk)
        return "Yes" if y is True else "No" if y is False else ""
    return ""


def usable_value(candidate: str, kind: str) -> bool:
    c = normalize_text(candidate)
    if not c or len(c) > 220:
        return False
    if re.fullmatch(r"\d+(?:\.\d+){1,5}\.?", c):
        return False
    if kind == "date":
        return parse_date(c) is not None
    if kind == "number":
        return parse_number(c) is not None
    if kind == "yesno":
        return parse_yesno(c) is not None
    if kind == "presence":
        return True
    # reject if still looks like a label only
    if c.endswith("?"):
        return False
    return True


def extract_doc_fields(doc: DocPack, specs: Dict[str, Dict[str, Any]]) -> Dict[str, str]:
    fields = dict(doc.fields or {})
    qblocks = build_question_blocks(doc.text)
    doc.qmap = qblocks
    # HVPQ / PIQ standard exact qid helpers first
    # Save qblocks in text-like fields for debug but not displayed.
    for key, spec in specs.items():
        if key in fields and fields[key]:
            continue
        val = extract_near_alias(doc.text, spec.get("aliases", []), spec.get("kind", "text"))
        if val:
            fields[key] = val
    # Specific smarter extractions
    fields.update(extract_specifics(doc.text, doc.kind))
    return fields


def extract_specifics(text: str, kind: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    blocks = build_question_blocks(text)
    # HVPQ known qids
    qid_map = {
        "vessel_name": ["1.1.2"], "imo_number": ["1.1.2"], "document_date": ["1.1.1"],
        "vessel_type": ["1.1.8", "1.1.9", "2.1.4"], "class_society": ["1.5.1"],
        "last_drydock": ["1.5.4"], "next_drydock_due": ["1.5.4"],
        "last_iws": ["1.5.5"], "next_iws_due": ["1.5.5"],
        "last_special_survey": ["1.5.6"], "next_special_survey_due": ["1.5.6"],
        "last_annual_survey": ["1.5.11"], "last_intermediate_survey": ["1.5.12"],
        "conditions_of_class": ["1.5.14"], "memoranda_of_class": ["1.5.16"], "flag_dispensation": ["1.5.18"],
        "last_psc_date": ["1.9.8"], "last_psc_port": ["1.9.8"], "psc_detained": ["1.9.8"],
        "incident_other": ["1.9.5"], "eedi_rating": ["1.2.1"], "eexi_rating": ["1.2.2"], "cii_rating": ["1.2.3"],
        "foam_type": ["5.3.1"], "foam_test_date": ["5.3.1"], "overboard_blanks": ["6.1.10"],
        "cargo_pressure_test": ["6.1.13"], "bunker_pressure_test": ["6.1.14"], "cargo_pressure": ["6.1.13"], "bunker_pressure": ["6.1.14"],
        "msmp": ["10.1.2"], "lmp": ["10.1.2"], "sdmbl": ["10.1.2"], "brake_testing_equipment": ["10.1.3"], "brake_test_date": ["10.1.4"],
        "split_drum": ["10.1.4", "10.1.3"], "mooring_diagram": ["10.1.3", "10.2.1", "10.7.1"], "manifold_diagram": ["10.8.1"],
        "crane_test_date": ["10.9.1"], "generator_count": ["11.3.1"],
    }
    for field, qids in qid_map.items():
        chunks = [blocks[q] for q in qids if q in blocks]
        if not chunks:
            continue
        spec = FIELD_SPECS.get(field, {})
        val = extract_near_alias("\n".join(chunks), spec.get("aliases", []), spec.get("kind", "text"))
        if val:
            out[field] = val
    # PIQ specific qids / unique patterns
    if "Technical Superintendent inspection completed" in text:
        out["technical_superintendent_section"] = get_context(text, "Technical Superintendent inspection completed", 16)
    if "Marine Superintendent inspection completed" in text:
        out["marine_superintendent_section"] = get_context(text, "Marine Superintendent inspection completed", 12)
    if "Date of PSC" in text or "Port State Control" in text:
        out["psc_section"] = get_context(text, "Date of PSC", 12) or get_context(text, "Port State Control", 12)
    # incidents section
    if "5.7." in text or "incidents occurred during" in text:
        out["piq_incident_section"] = get_context(text, "5.7. Safety Management", 80) or get_context(text, "Have any of the following incidents occurred", 80)
    # PIQ tank inspection dates
    for k, label in [
        ("cargo_tank_oldest_inspection", "oldest inspection report for all cargo and slop"),
        ("ballast_tank_oldest_inspection", "oldest inspection report for all ballast"),
        ("void_space_oldest_inspection", "oldest inspection report for all void"),
    ]:
        d = parse_date(get_context(text, label, 3))
        if d:
            out[k] = d.isoformat()
    return out


def get_context(text: str, needle: str, n_lines: int = 10) -> str:
    lines = text.splitlines()
    cneedle = clean_for_compare(needle)
    for i, line in enumerate(lines):
        if cneedle in clean_for_compare(line):
            return "\n".join(lines[i:i+n_lines])
    return ""

# -----------------------------
# Observation library parsing
# -----------------------------

def parse_observation_workbook(uploaded) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    if uploaded is None:
        return pd.DataFrame(), {}
    try:
        b = uploaded.getvalue()
        df = pd.read_excel(io.BytesIO(b), header=None, dtype=str).fillna("")
        rows = []
        current_question = ""
        for _, r in df.iterrows():
            cells = [normalize_text(x) for x in r.tolist()]
            line = " ".join([x for x in cells if x])
            m = re.search(r"Observations for Question:\s*([\d.]+)", line, flags=re.I)
            if m:
                current_question = m.group(1).rstrip(".")
                continue
            if len(cells) >= 2 and cells[0].lower() in {"s.no", "sno", "s no"}:
                continue
            if len(cells) >= 2 and cells[1]:
                rows.append({"source_question": current_question, "sno": cells[0], "observation": cells[1]})
            elif line and not line.lower().startswith("observations"):
                rows.append({"source_question": current_question, "sno": "", "observation": line})
        out = pd.DataFrame(rows)
        if out.empty:
            return out, {}
        out["question_refs"] = out["observation"].apply(extract_question_refs)
        out["category"] = out["observation"].apply(categorize_observation)
        summary = {
            "count": int(len(out)),
            "category_counts": out["category"].value_counts().to_dict(),
            "top_question_refs": count_question_refs(out).head(20).to_dict("records"),
        }
        return out, summary
    except Exception as e:
        return pd.DataFrame(), {"error": str(e)}


def extract_question_refs(s: str) -> List[str]:
    refs = re.findall(r"\b\d{1,2}\s*\.\s*\d+(?:\s*\.\s*\d+){0,3}\b", s)
    out = []
    for r in refs:
        r = re.sub(r"\s+", "", r).strip(".")
        # avoid dates like 03.10.2024 - crude filter
        parts = r.split(".")
        if len(parts) >= 2 and not (len(parts) == 3 and len(parts[-1]) == 4):
            out.append(r)
    return sorted(set(out))


def categorize_observation(s: str) -> str:
    cs = clean_for_compare(s)
    for area, _, keywords, _ in TARGETED_CHECKS:
        if any(clean_for_compare(k) in cs for k in keywords):
            return area
    return "General HVPQ accuracy"


def count_question_refs(df: pd.DataFrame) -> pd.DataFrame:
    counts: Dict[str, int] = {}
    for refs in df.get("question_refs", []):
        for r in refs:
            counts[r] = counts.get(r, 0) + 1
    return pd.DataFrame([{"question_ref": k, "count": v} for k, v in sorted(counts.items(), key=lambda x: -x[1])])

# -----------------------------
# Rules
# -----------------------------

def run_rules(hvpq: DocPack, piq: DocPack, class_doc: DocPack, q88: DocPack, xml_doc: DocPack, obs_df: pd.DataFrame, obs_summary: Dict[str, Any], settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    docs = {"HVPQ": hvpq, "PIQ": piq, "Class": class_doc, "Q88": q88, "XML": xml_doc}
    for d in docs.values():
        if d and (d.text or d.fields):
            d.fields.update(extract_doc_fields(d, FIELD_SPECS))

    findings.extend(rule_xml_diagnostic(xml_doc))
    findings.extend(rule_cross_doc_mismatches(hvpq, piq, class_doc, q88, xml_doc, settings))
    findings.extend(rule_incident_declaration(hvpq, piq, settings))
    findings.extend(rule_superintendent_gaps(piq, settings))
    findings.extend(rule_due_dates(hvpq, piq, class_doc, q88, settings))
    findings.extend(rule_targeted_ship_checks(hvpq, piq, class_doc, q88, obs_summary, settings))
    findings.extend(rule_observation_library_coverage(hvpq, piq, class_doc, q88, obs_df, obs_summary, settings))
    return findings


def rule_xml_diagnostic(xml_doc: DocPack) -> List[Finding]:
    if not xml_doc or not xml_doc.meta:
        return []
    warning = xml_doc.meta.get("xml_warning", "")
    if warning and "Control IDs need" in warning:
        return [Finding(
            area="XML / JSON Input",
            check="HVPQ XML response mapping",
            status="MANUAL CHECK",
            risk="MEDIUM",
            xml_value=f"Vessel={xml_doc.meta.get('vessel_name','')}; IMO={xml_doc.meta.get('imo_number','')}; Template={xml_doc.meta.get('template_variant','')}; Responses={xml_doc.meta.get('response_count','')}",
            reason="The uploaded OCIMF response XML provides reliable raw answers but uses control GUIDs. It does not expose human-readable HVPQ question labels unless a control mapping/template file is supplied.",
            required_action="For highest accuracy, upload/export the HVPQ as JSON/Excel with question text, or provide a control-map file with columns ctrl, qid, label. The app will still use the PDF for labels and XML for document identity/metadata.",
            source=xml_doc.name
        )]
    return []


def value_from(doc: DocPack, field: str) -> str:
    if not doc:
        return ""
    v = doc.fields.get(field, "") if doc.fields else ""
    if v is None:
        return ""
    return normalize_text(v)


def normalize_for_kind(value: str, kind: str) -> str:
    if kind == "date":
        d = parse_date(value)
        return d.isoformat() if d else ""
    if kind == "number":
        n = parse_number(value)
        return f"{n:g}" if n is not None else ""
    if kind == "yesno":
        y = parse_yesno(value)
        return "Yes" if y is True else "No" if y is False else ""
    return normalize_text(value)


def compare_values(a: str, b: str, kind: str, loose: bool = False) -> Tuple[bool, str]:
    aa, bb = normalize_for_kind(a, kind), normalize_for_kind(b, kind)
    if not aa or not bb:
        return False, "missing value"
    if kind == "date":
        return aa == bb, f"date compare {aa} vs {bb}"
    if kind == "number":
        try:
            return abs(float(aa) - float(bb)) < 0.01, f"number compare {aa} vs {bb}"
        except Exception:
            return aa == bb, f"number text compare {aa} vs {bb}"
    if kind == "yesno":
        return aa == bb, f"yes/no compare {aa} vs {bb}"
    ok = similar_enough(aa, bb, kind, loose=loose)
    return ok, f"semantic compare {aa} vs {bb}"


def rule_cross_doc_mismatches(hvpq: DocPack, piq: DocPack, class_doc: DocPack, q88: DocPack, xml_doc: DocPack, settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    # Compare HVPQ to PIQ/Class/Q88/XML for fields where values exist in both.
    for field, spec in FIELD_SPECS.items():
        kind = spec.get("kind", "text")
        risk = spec.get("risk", "MEDIUM")
        area = spec.get("area", "General")
        loose = bool(spec.get("loose", False))
        hv = value_from(hvpq, field)
        pv = value_from(piq, field)
        cv = value_from(class_doc, field)
        qv = value_from(q88, field)
        xv = value_from(xml_doc, field)
        pairs = []
        if hv and pv: pairs.append(("PIQ", pv))
        if hv and cv: pairs.append(("Class Status", cv))
        if hv and qv: pairs.append(("Q88", qv))
        if hv and xv: pairs.append(("HVPQ XML", xv))
        for source_name, other in pairs:
            ok, detail = compare_values(hv, other, kind, loose=loose)
            if not ok:
                # Don't flag missing via compare; missing handled elsewhere.
                if detail == "missing value":
                    continue
                findings.append(Finding(
                    area=area,
                    check=f"{field.replace('_',' ').title()} mismatch: HVPQ vs {source_name}",
                    status="MISMATCH",
                    risk=risk,
                    hvpq_value=hv,
                    piq_value=other if source_name == "PIQ" else "",
                    class_value=other if source_name == "Class Status" else "",
                    q88_value=other if source_name == "Q88" else "",
                    xml_value=other if source_name == "HVPQ XML" else "",
                    reason=f"Values are not equivalent after normalization ({detail}).",
                    required_action=f"Verify {field.replace('_',' ')} from source documents and correct the wrong declaration before submission.",
                    source=f"HVPQ vs {source_name}"
                ))
        # if Class/Q88 uploaded and HVPQ has important field but source cannot extract, add targeted manual check for high-risk only
        if risk in {"CRITICAL", "HIGH"} and hv:
            for source_name, doc in [("Class Status", class_doc), ("Q88", q88)]:
                if doc and doc.text and not value_from(doc, field) and area in {"Class / Survey", "Class", "Certificates", "PSC"}:
                    findings.append(Finding(
                        area=area,
                        check=f"{field.replace('_',' ').title()} requires {source_name} verification",
                        status="MANUAL CHECK",
                        risk="MEDIUM",
                        hvpq_value=hv,
                        reason=f"{source_name} was uploaded but the app could not confidently extract this field due to variable format.",
                        required_action=f"Ship/office to verify HVPQ value against {source_name}. If correct, mark verified; if not, update HVPQ.",
                        source=source_name
                    ))
    return dedupe_findings(findings)


def rule_incident_declaration(hvpq: DocPack, piq: DocPack, settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    hv_inc = value_from(hvpq, "incident_other")
    hv_section = "\n".join([hvpq.qmap.get("1.9.5", ""), hvpq.qmap.get("1.9.6", "")]) if hvpq and hvpq.qmap else ""
    piq_sec = value_from(piq, "piq_incident_section")
    # Extract PIQ incident yes/no counts in 5.7 section
    piq_yes = len(re.findall(r"\bYes\b", piq_sec, flags=re.I)) if piq_sec else 0
    piq_no = len(re.findall(r"\bNo\b", piq_sec, flags=re.I)) if piq_sec else 0
    hv_yesno = parse_yesno(hv_inc or hv_section)

    if hv_yesno is False and (not piq_sec or piq_yes == 0):
        findings.append(Finding(
            area="Incidents",
            check="No incident declaration in HVPQ/PIQ",
            status="MANUAL CHECK",
            risk="MEDIUM",
            hvpq_value=hv_inc or "No / no incident details found",
            piq_value=f"PIQ 5.7 Yes count={piq_yes}, No count={piq_no}" if piq_sec else "PIQ incident section not clearly extracted",
            reason="No incident was detected in HVPQ/PIQ. This is not automatically wrong, but historic observations commonly arise when incidents existed but were not declared.",
            required_action="Send to vessel/office for positive confirmation: no pollution, grounding/touch bottom, blackout, main engine failure, mooring rope parting, injury/LTI, cargo/steering/critical equipment incident in previous 12 months.",
            source="HVPQ/PIQ"
        ))
    elif hv_yesno is True and piq_sec and piq_yes == 0:
        findings.append(Finding(
            area="Incidents",
            check="HVPQ lists incident but PIQ incident questions appear all No",
            status="MISMATCH",
            risk="HIGH",
            hvpq_value=shorten(hv_section, 300),
            piq_value=f"PIQ 5.7 Yes count={piq_yes}, No count={piq_no}",
            reason="HVPQ appears to declare an incident, while PIQ incident checklist appears to have no positive incident declaration.",
            required_action="Review PIQ 5.7 answers against HVPQ 1.9.5/1.9.6 and update if the incident falls under any PIQ incident category.",
            source="HVPQ vs PIQ",
            question_ref="HVPQ 1.9.5/1.9.6; PIQ 5.7"
        ))
    elif not hv_inc and not hv_section:
        findings.append(Finding(
            area="Incidents",
            check="HVPQ incident section not confidently extracted",
            status="MANUAL CHECK",
            risk="MEDIUM",
            reason="Could not confidently extract HVPQ 1.9.5/1.9.6 incident declaration.",
            required_action="Verify HVPQ 1.9.5/1.9.6 manually and confirm consistency with PIQ 5.7.",
            source="HVPQ"
        ))
    return findings


def rule_superintendent_gaps(piq: DocPack, settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    if not piq or not piq.text:
        return findings
    piq_date = parse_date(value_from(piq, "document_date")) or date.today()
    for label, section_key, max_months in [
        ("Technical Superintendent", "technical_superintendent_section", float(settings.get("tech_gap_months", 7.0))),
        ("Marine Superintendent", "marine_superintendent_section", float(settings.get("marine_gap_months", 12.0))),
    ]:
        sec = value_from(piq, section_key)
        if not sec:
            findings.append(Finding(
                area="Management Oversight",
                check=f"{label} visit section extraction",
                status="MANUAL CHECK",
                risk="MEDIUM",
                reason=f"Could not confidently extract {label} visit history from PIQ.",
                required_action=f"Verify {label} inspection dates and gap rule manually.",
                source="PIQ"
            ))
            continue
        dates = sorted(set(d for d in extract_dates(sec)))
        if not dates:
            findings.append(Finding(
                area="Management Oversight",
                check=f"{label} visit dates missing",
                status="MANUAL CHECK",
                risk="HIGH",
                piq_value=shorten(sec, 250),
                reason=f"{label} section found, but dates were not extracted.",
                required_action=f"Confirm {label} last/second/third inspection dates.",
                source="PIQ"
            ))
            continue
        # Gaps between successive dates and from latest date to PIQ date
        all_dates = dates + [piq_date]
        for a, b in zip(all_dates, all_dates[1:]):
            days = (b - a).days
            allowed = approx_months_to_days(max_months)
            if days > allowed:
                findings.append(Finding(
                    area="Management Oversight",
                    check=f"{label} inspection gap exceeds {max_months:g} months",
                    status="NON-COMPLIANT",
                    risk="CRITICAL",
                    piq_value=f"{a.isoformat()} to {b.isoformat()} = {days} days",
                    reason=f"Configured strict maximum gap is {max_months:g} months. Any exceedance is flagged.",
                    required_action=f"Provide office justification if applicable and schedule/record compliant {label} inspection. Update PIQ if dates are wrong.",
                    source="PIQ"
                ))
    return findings


def extract_dates(text: str) -> List[date]:
    dates = []
    # Find robust dates; avoid question numbers.
    patterns = [
        r"\b\d{1,2}\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(?:20\d{2}|19\d{2})\b",
        r"\b(?:20\d{2}|19\d{2})[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}[-/](?:20\d{2}|19\d{2})\b",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.I):
            d = parse_date(m.group(0))
            if d:
                dates.append(d)
    return dates


def approx_months_to_days(months: float) -> int:
    return int(round(months * 30.4375))


def rule_due_dates(hvpq: DocPack, piq: DocPack, class_doc: DocPack, q88: DocPack, settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    ref_date = parse_date(value_from(hvpq, "document_date")) or parse_date(value_from(piq, "document_date")) or date.today()
    warning_days = int(settings.get("due_warning_days", 180))
    cert_warning_days = int(settings.get("cert_warning_days", 90))
    # due survey checks
    for field, label in [("next_drydock_due", "Next drydock"), ("next_iws_due", "Next IWS"), ("next_special_survey_due", "Next special survey")]:
        d = parse_date(value_from(hvpq, field))
        if d:
            days = (d - ref_date).days
            if days < 0:
                findings.append(Finding("Class / Survey", label, "OVERDUE", "CRITICAL", hvpq_value=d.isoformat(), reason=f"Due date is {abs(days)} days before document/reference date {ref_date}.", required_action="Verify latest class status/certificate and update HVPQ immediately.", source="HVPQ"))
            elif days <= warning_days:
                findings.append(Finding("Class / Survey", label, "DUE SOON", "MEDIUM", hvpq_value=d.isoformat(), reason=f"Due within {days} days of reference date {ref_date}.", required_action="Verify class status and ensure no outdated HVPQ entry.", source="HVPQ"))
    # Tank inspection due based on oldest inspection + frequency
    freq_months = int(settings.get("tank_frequency_months", 12))
    tank_warning = int(settings.get("tank_warning_days", 60))
    for field, label in [("cargo_tank_oldest_inspection", "Cargo/slop tank inspection"), ("ballast_tank_oldest_inspection", "Ballast tank inspection"), ("void_space_oldest_inspection", "Void space inspection")]:
        d = parse_date(value_from(piq, field) or value_from(hvpq, field))
        if d:
            due = d + relativedelta(months=freq_months)
            days = (due - ref_date).days
            if days < 0:
                findings.append(Finding("Tank Inspection", label, "OVERDUE", "HIGH", hvpq_value=value_from(hvpq, field), piq_value=value_from(piq, field), reason=f"Oldest inspection {d}; due {due}; overdue by {abs(days)} days.", required_action="Verify tank inspection sequence and update PIQ/HVPQ if wrong.", source="PIQ/HVPQ"))
            elif days <= tank_warning:
                findings.append(Finding("Tank Inspection", label, "DUE SOON", "MEDIUM", hvpq_value=value_from(hvpq, field), piq_value=value_from(piq, field), reason=f"Oldest inspection {d}; due {due}; due within {days} days.", required_action="Ship to confirm next inspection plan and whether PIQ/HVPQ dates are correct.", source="PIQ/HVPQ"))
    return findings


def rule_targeted_ship_checks(hvpq: DocPack, piq: DocPack, class_doc: DocPack, q88: DocPack, obs_summary: Dict[str, Any], settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    combined_text = "\n".join([d.text for d in [hvpq, piq, class_doc, q88] if d and d.text])
    obs_cats = set((obs_summary or {}).get("category_counts", {}).keys())
    # Only show targeted checks if relevant docs are missing or extract confidence low. Not pass checks.
    for area, check, keywords, action in TARGETED_CHECKS:
        # Only if observation library indicates this area OR the document contains the area keywords.
        relevant = area in obs_cats or any(clean_for_compare(k) in clean_for_compare(combined_text) for k in keywords)
        if not relevant:
            continue
        # If no concrete mismatch already, create manual targeted check only where evidence doc isn't enough.
        # This is intended for ship export.
        missing_sources = []
        if area in {"Class / Certificates", "Certificates", "Class / Survey"} and not (class_doc and class_doc.text):
            missing_sources.append("Class Status")
        if area in {"Mooring", "Diagrams", "Pollution / Cargo", "Lifting Gear", "Cargo / IGS"} and not (q88 and q88.text):
            # Q88 not mandatory but useful. Avoid too much noise by lower risk.
            pass
        findings.append(Finding(
            area=area,
            check=check,
            status="TARGETED CHECK",
            risk="MEDIUM",
            reason="This topic appears in the uploaded documents and/or historical observation library. It is a recurring HVPQ/PIQ observation area and should be ship-confirmed even where extraction is uncertain.",
            required_action=action,
            source="Observation library + uploaded docs"
        ))
    return dedupe_findings(findings)


def rule_observation_library_coverage(hvpq: DocPack, piq: DocPack, class_doc: DocPack, q88: DocPack, obs_df: pd.DataFrame, obs_summary: Dict[str, Any], settings: Dict[str, Any]) -> List[Finding]:
    findings: List[Finding] = []
    if obs_df is None or obs_df.empty:
        return findings
    # Generate a concise library-driven priority note, not 213 rows of noise.
    cat_counts = obs_summary.get("category_counts", {}) if obs_summary else {}
    if cat_counts:
        top = ", ".join([f"{k}({v})" for k, v in list(cat_counts.items())[:8]])
        findings.append(Finding(
            area="Observation Library",
            check="Historical observation library applied",
            status="INFO",
            risk="INFO",
            reason=f"Loaded {len(obs_df)} historical observations. Top risk families: {top}.",
            required_action="Use the targeted checks tab/export as the ship-facing pre-inspection verification list.",
            source="Observation Excel"
        ))
    # If high-frequency qrefs exist but not found in HVPQ text, flag as manual because it may be missing/extraction issue.
    qcounts = count_question_refs(obs_df)
    if not qcounts.empty and hvpq and hvpq.text:
        for _, row in qcounts.head(20).iterrows():
            qid = row["question_ref"]
            if row["count"] >= int(settings.get("obs_qid_threshold", 2)):
                if qid not in hvpq.qmap and qid not in hvpq.text:
                    findings.append(Finding(
                        area="Observation Library",
                        check=f"Recurring observed HVPQ question {qid} not clearly extracted",
                        status="MANUAL CHECK",
                        risk="MEDIUM",
                        reason=f"Question {qid} appears {row['count']} times in the historical observation library, but was not confidently extracted from current HVPQ PDF.",
                        required_action=f"Verify HVPQ question {qid} manually if applicable to this vessel.",
                        source="Observation Excel",
                        question_ref=qid
                    ))
    return findings

# -----------------------------
# Output / UI helpers
# -----------------------------

def shorten(s: Any, n: int = 180) -> str:
    s = normalize_text(s)
    return s if len(s) <= n else s[:n-3] + "..."


def dedupe_findings(findings: List[Finding]) -> List[Finding]:
    seen = set()
    out = []
    for f in findings:
        key = (f.area, f.check, f.status, f.hvpq_value, f.piq_value, f.class_value, f.q88_value)
        if key not in seen:
            seen.add(key)
            out.append(f)
    return out


def findings_to_df(findings: List[Finding], show_info: bool = False) -> pd.DataFrame:
    rows = [asdict(f) for f in findings]
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    if not show_info:
        df = df[~df["risk"].isin(["PASS", "INFO", "LOW"])]
    df["risk_rank"] = df["risk"].map(RISK_ORDER).fillna(2)
    df = df.sort_values(["risk_rank", "area", "check"], ascending=[False, True, True]).drop(columns=["risk_rank"])
    return df


def to_excel_bytes(df: pd.DataFrame, obs_df: pd.DataFrame = None, debug_fields: Dict[str, Any] = None) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if df.empty:
            pd.DataFrame([{"Message": "No actionable findings generated."}]).to_excel(writer, index=False, sheet_name="Findings")
        else:
            df.to_excel(writer, index=False, sheet_name="Findings")
        if obs_df is not None and not obs_df.empty:
            obs_df.head(500).to_excel(writer, index=False, sheet_name="Obs Library")
        if debug_fields:
            rows = []
            for doc, fields in debug_fields.items():
                for k, v in fields.items():
                    if k.startswith("_"):
                        continue
                    rows.append({"Document": doc, "Field": k, "Value": shorten(v, 500)})
            pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Extracted Fields")
        # basic formatting
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = 10
                col_letter = col[0].column_letter
                for cell in col[:1000]:
                    try:
                        max_len = max(max_len, min(len(str(cell.value or "")), 60))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = max_len + 2
    return output.getvalue()

# -----------------------------
# Streamlit app
# -----------------------------

def main():
    st.set_page_config(page_title="HVPQ / PIQ Vetting Checker", layout="wide")
    st.title("HVPQ / PIQ Vetting Observation Checker v4")
    st.caption("Mismatch-focused. Observation-library driven. Designed to export simple targeted checks for ship/office verification.")

    with st.sidebar:
        st.header("Upload documents")
        hvpq_pdf = st.file_uploader("HVPQ PDF", type=["pdf"], key="hvpq_pdf")
        hvpq_xml = st.file_uploader("HVPQ XML / response XML (optional)", type=["xml"], key="hvpq_xml")
        ctrl_map_file = st.file_uploader("Optional XML control map (ctrl, qid, label)", type=["xlsx", "xls", "csv"], key="ctrl_map")
        piq_pdf = st.file_uploader("PIQ PDF", type=["pdf"], key="piq_pdf")
        class_file = st.file_uploader("Class Status / Certificate Status (PDF/XLSX/CSV/TXT)", type=["pdf", "xlsx", "xls", "csv", "txt"], key="class")
        q88_file = st.file_uploader("Q88 / Vessel Particulars (optional)", type=["pdf", "xlsx", "xls", "csv", "txt"], key="q88")
        st.divider()
        obs_hvpq = st.file_uploader("HVPQ Observation Library Excel", type=["xlsx", "xls"], key="obs_hvpq")
        obs_inc = st.file_uploader("Incident Observation Library Excel", type=["xlsx", "xls"], key="obs_inc")
        st.divider()
        st.header("Rules")
        tech_gap = st.number_input("Technical Superintendent max gap (months)", value=7.0, min_value=1.0, step=0.5)
        marine_gap = st.number_input("Marine Superintendent max gap (months)", value=12.0, min_value=1.0, step=0.5)
        due_warning = st.number_input("Survey due warning days", value=180, min_value=0, step=30)
        cert_warning = st.number_input("Certificate expiry warning days", value=90, min_value=0, step=30)
        tank_warning = st.number_input("Tank inspection due warning days", value=60, min_value=0, step=15)
        show_info = st.checkbox("Show INFO / PASS / debug rows", value=False)
        run = st.button("Run verification", type="primary")

    if not run:
        st.info("Upload HVPQ + PIQ and optionally Class Status/Q88/observation files, then run verification.")
        st.markdown("""
        **Important design point:** this app does not try to certify the whole HVPQ. It finds likely mismatches and creates ship-facing targeted verification items based on historical observation patterns.
        """)
        return

    settings = {
        "tech_gap_months": tech_gap,
        "marine_gap_months": marine_gap,
        "due_warning_days": due_warning,
        "cert_warning_days": cert_warning,
        "tank_warning_days": tank_warning,
        "tank_frequency_months": 12,
        "obs_qid_threshold": 2,
    }

    # Optional control map for XML
    ctrl_map = None
    if ctrl_map_file:
        try:
            b = ctrl_map_file.getvalue()
            if ctrl_map_file.name.lower().endswith(".csv"):
                ctrl_map = pd.read_csv(io.BytesIO(b), dtype=str)
            else:
                ctrl_map = pd.read_excel(io.BytesIO(b), dtype=str)
        except Exception as e:
            st.warning(f"Could not read control map: {e}")

    with st.spinner("Extracting documents..."):
        hvpq = read_any_upload(hvpq_pdf, "HVPQ") if hvpq_pdf else DocPack("", "HVPQ")
        piq = read_any_upload(piq_pdf, "PIQ") if piq_pdf else DocPack("", "PIQ")
        class_doc = read_any_upload(class_file, "Class Status") if class_file else DocPack("", "Class Status")
        q88 = read_any_upload(q88_file, "Q88") if q88_file else DocPack("", "Q88")
        xml_doc = DocPack("", "HVPQ XML")
        if hvpq_xml:
            fields, meta = parse_hvpq_xml(hvpq_xml.getvalue(), ctrl_map)
            xml_doc = DocPack(hvpq_xml.name, "HVPQ XML", xml_to_searchable_text(fields, meta), fields, meta=meta)
        obs_frames = []
        obs_summaries = []
        for u in [obs_hvpq, obs_inc]:
            df, summ = parse_observation_workbook(u)
            if not df.empty:
                df["library_file"] = u.name
                obs_frames.append(df)
                obs_summaries.append(summ)
        obs_df = pd.concat(obs_frames, ignore_index=True) if obs_frames else pd.DataFrame()
        obs_summary = merge_obs_summaries(obs_df)

    with st.spinner("Running rules..."):
        findings = run_rules(hvpq, piq, class_doc, q88, xml_doc, obs_df, obs_summary, settings)
        df = findings_to_df(findings, show_info=show_info)

    # Dashboard
    c1, c2, c3, c4, c5 = st.columns(5)
    if not df.empty:
        c1.metric("Actionable rows", len(df))
        c2.metric("Critical", int((df["risk"] == "CRITICAL").sum()))
        c3.metric("High", int((df["risk"] == "HIGH").sum()))
        c4.metric("Medium", int((df["risk"] == "MEDIUM").sum()))
        c5.metric("Manual/Targeted", int(df["status"].isin(["MANUAL CHECK", "TARGETED CHECK"]).sum()))
    else:
        c1.metric("Actionable rows", 0)

    tabs = st.tabs(["Findings", "Targeted ship checklist", "Observation library", "Extracted fields", "XML diagnostic"])
    with tabs[0]:
        st.subheader("Actionable mismatch / manual-check register")
        if df.empty:
            st.success("No actionable findings generated from the extracted data. This does not mean the HVPQ is fully verified; use targeted checks if required.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            excel = to_excel_bytes(df, obs_df, debug_fields={
                "HVPQ": hvpq.fields, "PIQ": piq.fields, "Class": class_doc.fields, "Q88": q88.fields, "XML": xml_doc.fields
            })
            st.download_button("Download Excel register", excel, file_name="hvpq_piq_verification_register.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("Download CSV register", df.to_csv(index=False).encode("utf-8"), file_name="hvpq_piq_verification_register.csv", mime="text/csv")
    with tabs[1]:
        ship_df = df[df["status"].isin(["TARGETED CHECK", "MANUAL CHECK", "MISMATCH", "NON-COMPLIANT", "OVERDUE", "DUE SOON"])] if not df.empty else df
        st.subheader("Ship-facing targeted checks")
        st.caption("This is the simplified list you can export/send to vessel. It intentionally avoids PASS rows.")
        st.dataframe(ship_df[[c for c in ["area","check","risk","hvpq_value","piq_value","class_value","q88_value","reason","required_action","question_ref"] if c in ship_df.columns]], use_container_width=True, hide_index=True)
    with tabs[2]:
        st.subheader("Observation library analysis")
        if obs_df.empty:
            st.warning("No observation library loaded. Upload your HVPQ/Incident observation Excel files for better targeted checks.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                st.write("Category counts")
                st.dataframe(pd.DataFrame([{"category": k, "count": v} for k, v in obs_summary.get("category_counts", {}).items()]), hide_index=True)
            with col2:
                st.write("Top question references")
                st.dataframe(pd.DataFrame(obs_summary.get("top_question_refs", [])), hide_index=True)
            st.write("Observation rows")
            st.dataframe(obs_df, use_container_width=True, hide_index=True)
    with tabs[3]:
        st.subheader("Extracted fields / debug")
        for name, doc in [("HVPQ", hvpq), ("PIQ", piq), ("Class", class_doc), ("Q88", q88), ("XML", xml_doc)]:
            with st.expander(name, expanded=False):
                if doc and doc.fields:
                    st.json({k: shorten(v, 500) for k, v in doc.fields.items() if not k.startswith("_")})
                else:
                    st.write("No fields extracted.")
                st.caption(f"Text length: {len(doc.text) if doc else 0}")
    with tabs[4]:
        st.subheader("XML / JSON guidance")
        if xml_doc and xml_doc.meta:
            st.json(xml_doc.meta)
        st.markdown("""
        **Best input hierarchy for HVPQ accuracy:**
        1. JSON/Excel export with `question id + question text + answer` — best.
        2. OCIMF XML response **plus** control map `ctrl -> qid/label` — very good.
        3. HVPQ PDF only — workable, but extraction can be noisy because tables split across lines.
        
        The response XML alone helps with vessel identity and raw answer values, but the control GUIDs are not useful for question-level validation unless mapped to the HVPQ template.
        """)


def merge_obs_summaries(obs_df: pd.DataFrame) -> Dict[str, Any]:
    if obs_df is None or obs_df.empty:
        return {}
    return {
        "count": int(len(obs_df)),
        "category_counts": obs_df["category"].value_counts().to_dict(),
        "top_question_refs": count_question_refs(obs_df).head(30).to_dict("records"),
    }


if __name__ == "__main__":
    main()
