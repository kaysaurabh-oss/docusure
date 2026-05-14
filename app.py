from __future__ import annotations

import io, re, zipfile, xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict
from datetime import date, datetime
from difflib import SequenceMatcher
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import streamlit as st
from dateutil import parser as dateparser
from dateutil.relativedelta import relativedelta

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

APP_VERSION = "v7"

# -------------------------
# Data models
# -------------------------
@dataclass
class Field:
    value: str = ""
    confidence: str = ""
    source: str = ""
    method: str = ""
    evidence: str = ""

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
    evidence: str = ""

PASS_STATUSES = {"PASS", "INFO"}
ACTION_STATUSES = {"MISMATCH", "MANUAL CHECK", "MISSING", "WARNING"}

# Month-name based date patterns only. This intentionally avoids interpreting qids like 1.5.11 as dates.
DATE_RE = re.compile(
    r"(?:\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|"
    r"[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}|"
    r"\d{1,2}[./-][A-Za-z]{3,9}[./-]\d{4})",
    re.I,
)

CERT_KEYWORDS = {
    "safety_equipment_expiry": ["Safety Equipment Certificate", "Safety Equipment", "SEC"],
    "safety_radio_expiry": ["Safety Radio Certificate", "Safety Radio", "SRC"],
    "safety_construction_expiry": ["Safety Construction Certificate", "Safety Construction", "SCC"],
    "loadline_expiry": ["Loadline Certificate", "Load Line", "International Loadline Certificate", "ILC"],
    "iopp_expiry": ["International Oil Pollution Prevention Certificate", "IOPPC", "OPP (MARPOL Annex I)"],
    "ibwmc_expiry": ["International Ballast Water Management Certificate", "IBWMC", "BWM"],
    "smc_expiry": ["Safety Management Certificate", "ISM Safety Management Certificate", "SMC"],
    "issc_expiry": ["International Ship Security Certificate", "ISSC"],
    "doc_expiry": ["Document of Compliance", "DOC"],
    "cof_chem_expiry": ["Certificate of Fitness (COF)", "Certificate of Fitness (Expiry", "Certificate of Fitness", "Chemicals in Bulk"],
    "class_certificate_expiry": ["Class Certificate", "Certificate of Class (COC)", "Certificate of Class"],
    "uscg_coc_expiry": ["USCG Certificate of Compliance", "USCGCOC", "USCG Certificate"],
    "cofr_expiry": ["U.S. Certificate of Financial Responsibility", "COFR"],
    "clc_oil_expiry": ["Civil Liability Convention Certificate (1992)", "CLC 1992", "Civil Liability Certificates"],
    "bunker_clc_expiry": ["Civil Liability Convention 2001", "CLBC", "Bunker Oil Pollution"],
    "wreck_removal_expiry": ["Wreck removal Convention Certificate", "WRC", "Removal of Wrecks"],
}

# Meaningful cross-source comparison map. PIQ and Class Status are not forced into non-overlapping fields.
COMPARE_MAP = [
    ("vessel_name", "Identity", "Vessel name", ["hvpq", "piq", "class", "q88"]),
    ("imo_number", "Identity", "IMO number", ["hvpq", "class", "q88"]),
    ("flag", "General", "Flag", ["hvpq", "class", "q88"]),
    ("vessel_type", "General", "Vessel type", ["hvpq", "piq", "class", "q88"]),
    ("class_society", "Class", "Class society", ["hvpq", "class", "q88"]),
    ("class_notation", "Class", "Class notation", ["hvpq", "class", "q88"]),
    ("conditions_of_class", "Class / Survey", "Open conditions of class", ["hvpq", "class", "q88"]),
    ("memoranda_of_class", "Class / Survey", "Memoranda of class", ["hvpq", "class", "q88"]),
    ("flag_dispensation", "Class / Survey", "Flag dispensation", ["hvpq"]),
    ("last_drydock", "Class / Survey", "Last drydock", ["hvpq", "q88"]),
    ("next_drydock_due", "Class / Survey", "Next drydock / docking survey due", ["hvpq", "class", "q88"]),
    ("last_iws", "Class / Survey", "Last IWS / docking survey", ["hvpq", "class", "q88"]),
    ("next_iws_due", "Class / Survey", "Next IWS / docking survey due", ["hvpq", "class", "q88"]),
    ("last_special_survey", "Class / Survey", "Last special survey", ["hvpq", "class", "q88"]),
    ("next_special_survey_due", "Class / Survey", "Next special survey due", ["hvpq", "class", "q88"]),
    ("last_annual_survey", "Class / Survey", "Last annual survey", ["hvpq", "class"]),
    ("last_intermediate_survey", "Class / Survey", "Last intermediate survey", ["hvpq", "class"]),
    ("last_psc_date", "PSC", "Last PSC date", ["hvpq", "piq", "q88"]),
    ("last_psc_port", "PSC", "Last PSC port", ["hvpq", "piq", "q88"]),
    ("psc_detained", "PSC", "PSC detention status", ["hvpq", "piq"]),
    ("incident_pollution_grounding_collision", "Incidents", "Pollution/grounding/collision/allision incident", ["hvpq", "piq"]),
    ("incident_other", "Incidents", "Other incidents in past 12 months", ["hvpq", "piq"]),
    ("foam_type", "Firefighting", "Foam type", ["hvpq", "q88"]),
    ("foam_test_date", "Firefighting", "Foam test/supply date", ["hvpq", "q88"]),
    ("cargo_pressure", "Pollution / Cargo", "Cargo piping pressure test pressure", ["hvpq", "q88"]),
    ("bunker_pressure", "Pollution / Cargo", "Bunker piping pressure test pressure", ["hvpq", "q88"]),
    ("overboard_blanks", "Pollution / Cargo", "Overboard discharge blanks/testing", ["hvpq", "q88"]),
    ("sea_chest", "Pollution / Cargo", "Cargo sea chest / sea valves", ["hvpq", "q88"]),
    ("cargo_tank_oldest_inspection", "Tank Inspection", "Oldest cargo/slop tank inspection", ["piq"]),
    ("ballast_tank_oldest_inspection", "Tank Inspection", "Oldest ballast tank inspection", ["piq"]),
    ("void_oldest_inspection", "Tank Inspection", "Oldest void inspection", ["piq"]),
]

CERT_COMPARE_KEYS = [
    "safety_equipment_expiry", "safety_radio_expiry", "safety_construction_expiry", "loadline_expiry",
    "iopp_expiry", "ibwmc_expiry", "smc_expiry", "issc_expiry", "doc_expiry", "cof_chem_expiry",
    "class_certificate_expiry", "uscg_coc_expiry", "cofr_expiry", "clc_oil_expiry", "bunker_clc_expiry", "wreck_removal_expiry",
]

# -------------------------
# Low-level helpers
# -------------------------
def read_pdf(uploaded_file) -> str:
    if uploaded_file is None or fitz is None:
        return ""
    data = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file.getvalue()
    doc = fitz.open(stream=data, filetype="pdf")
    pages = []
    for i, page in enumerate(doc, start=1):
        pages.append(f"\n\n--- PAGE {i} ---\n" + page.get_text("text"))
    return "\n".join(pages)

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()

def clean_value(s: str) -> str:
    s = norm_space(s)
    s = re.sub(r"^[\s:;,.\-]+|[\s:;,.\-]+$", "", s)
    return s

def setf(d: Dict[str, Field], key: str, value: Any, source: str, method: str, evidence: str, confidence: str="HIGH"):
    if value is None:
        return
    value = str(value).strip()
    if not value:
        return
    if key not in d or not d[key].value or (d[key].confidence != "HIGH" and confidence == "HIGH"):
        d[key] = Field(clean_value(value), confidence, source, method, norm_space(evidence)[:1200])

def find_dates(text: str) -> List[date]:
    out = []
    for m in DATE_RE.finditer(text or ""):
        raw = m.group(0)
        try:
            out.append(dateparser.parse(raw, dayfirst=True).date())
        except Exception:
            try:
                out.append(dateparser.parse(raw.replace('.', ' '), dayfirst=True).date())
            except Exception:
                pass
    return out

def dstr(d: date) -> str:
    return d.isoformat() if isinstance(d, date) else str(d or "")

def extract_qblock(text: str, qid: str, max_chars: int=2500) -> str:
    # Start at exact question id, stop at next question id of similar numeric structure.
    pat = re.compile(rf"(?m)^\s*{re.escape(qid)}\b")
    m = pat.search(text or "")
    if not m:
        # fallback when page extraction keeps qid inside a line
        m = re.search(rf"{re.escape(qid)}\b", text or "")
    if not m:
        return ""
    start = m.start()
    rest = text[start:start+max_chars]
    nxt = re.search(r"(?m)^\s*\d+(?:\.\d+){1,4}\b", rest[5:])
    if nxt:
        rest = rest[:5+nxt.start()]
    return rest

def page_text(text: str, page_no: int) -> str:
    m = re.search(rf"--- PAGE {page_no} ---\n(.*?)(?=\n\n--- PAGE \d+ ---|\Z)", text, flags=re.S)
    return m.group(1) if m else ""

def answer_yesno(block: str) -> str:
    # last yes/no in a short block usually is the answer, but avoid child questions where possible.
    yn = re.findall(r"\b(Yes|No|Not applicable|N/A|NA)\b", block or "", flags=re.I)
    if yn:
        v = yn[-1]
        return "Not applicable" if v.lower().startswith("not") else ("Yes" if v.lower()=="yes" else "No" if v.lower()=="no" else v.upper())
    return ""

def after_label(block: str, label_regex: str) -> str:
    lines = [clean_value(x) for x in (block or "").splitlines() if clean_value(x)]
    joined = " ".join(lines)
    m = re.search(label_regex + r"\s*[:?]?\s*(.*?)(?:\s{2,}|$)", joined, flags=re.I)
    if m and m.group(1):
        return clean_value(m.group(1))
    for i, ln in enumerate(lines):
        if re.search(label_regex, ln, re.I):
            tail = re.sub(label_regex, "", ln, flags=re.I).strip(" :?-")
            if tail:
                return clean_value(tail)
            if i + 1 < len(lines):
                return clean_value(lines[i+1])
    return ""

def sim(a: str, b: str) -> float:
    return SequenceMatcher(None, norm_for_compare(a), norm_for_compare(b)).ratio()

def norm_for_compare(v: str) -> str:
    s = (v or "").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    replacements = {
        "nippon kaiji kyokai": "classnk",
        "class nk": "classnk",
        "nk ships": "classnk",
        "classnk": "classnk",
        "not applicable": "no",
        "n a": "no",
        "nil": "no",
        "none": "no",
        "double hull": "doublehull",
        "classification society": "",
        "1 classification society": "",
        "type of ship purpose intended service": "",
        "certificates sc se sf": "productchemical",
        "oil chemical carrier": "productchemical",
        "oil tanker chemical tanker": "productchemical",
        "product carrier": "productchemical",
        "products chemical tanker": "productchemical",
        "oil chemical carrier": "productchemical",
        "oil tanker chemical tanker": "productchemical",
        "other product carrier": "productchemical",
        "assuranceforeningen gard gjensidig": "gard",
    }
    s = norm_space(s)
    for k, val in replacements.items():
        s = s.replace(k, val)
    return norm_space(s)

def equivalent(a: str, b: str, field: str="") -> Tuple[bool, str]:
    a0, b0 = clean_value(a), clean_value(b)
    if not a0 or not b0:
        return True, "one or both values missing; no mismatch assumed"
    da, db = find_dates(a0), find_dates(b0)
    if re.search(r"date|expiry|survey|iws|drydock", field, re.I) or (da and db):
        if da and db:
            return da[0] == db[0], f"date compare {dstr(da[0])} vs {dstr(db[0])}"
    na, nb = norm_for_compare(a0), norm_for_compare(b0)
    if na == nb:
        return True, "normalized exact match"
    if {na, nb} <= {"yes", "no", "not applicable", "na"}:
        return na == nb, "yes/no compare"
    # for vessel names allow suffixes like LR2
    if field == "vessel_name":
        if na in nb or nb in na or sim(a0, b0) > 0.78:
            return True, "loose vessel-name match"
        return False, "material vessel-name mismatch"
    if field == "class_notation":
        # Do not over-flag notation formatting differences. Require only major feature disagreement.
        def has(tok, s): return tok in s
        major = ["egcs", "sox", "esp", "eedi", "csr"]
        disagreements = [t for t in major if has(t, na) != has(t, nb)]
        return not disagreements, "major class-notation token compare"
    if sim(a0, b0) >= 0.72:
        return True, f"semantic/fuzzy match {sim(a0,b0):.2f}"
    return False, f"values not equivalent after normalization ({na} vs {nb})"

# -------------------------
# Generic table parsers
# -------------------------
def cert_block(text: str) -> str:
    start = re.search(r"Date Issued\s+Date Expires|2\.\s*CERTIFICATES|Current Statutory Certificates", text or "", re.I)
    if not start:
        return ""
    end = re.search(r"\n\s*(Publications|2\.2\.1|3\.\s*CREW|Survey Status: Class|Documentation)\b", text[start.start():], re.I)
    return text[start.start(): start.start()+end.start()] if end else text[start.start():start.start()+8000]

def segment_between_keywords(block: str, start_patterns: List[str], all_patterns: List[List[str]]) -> str:
    if not block:
        return ""
    low = block.lower()
    positions = []
    for pats in all_patterns:
        best = None
        for p in pats:
            idx = low.find(p.lower())
            if idx >= 0:
                best = idx if best is None else min(best, idx)
        if best is not None:
            positions.append(best)
    starts = []
    for p in start_patterns:
        idx = low.find(p.lower())
        if idx >= 0:
            starts.append(idx)
    if not starts:
        return ""
    s = min(starts)
    next_positions = [p for p in positions if p > s]
    e = min(next_positions) if next_positions else min(len(block), s+800)
    return block[s:e]

def extract_hvpq_cert_dates(text: str, f: Dict[str, Field]):
    block = cert_block(text)
    allp = list(CERT_KEYWORDS.values())
    for key, pats in CERT_KEYWORDS.items():
        seg = segment_between_keywords(block, pats, allp)
        if not seg:
            continue
        dates = find_dates(seg)
        # HVPQ cert table order: Issued, Expires, Last Annual, Last Intermediate.
        # In Q88 order differs; this function is HVPQ only.
        if key == "cofr_expiry" and len(dates) >= 2:
            expiry = dates[1]
        elif len(dates) >= 2:
            expiry = dates[1]
        else:
            continue
        setf(f, key, dstr(expiry), "HVPQ", "certificate table row: expiry is 2nd date", seg)

def extract_q88_cert_dates(text: str, f: Dict[str, Field]):
    block = cert_block(text)
    allp = list(CERT_KEYWORDS.values())
    for key, pats in CERT_KEYWORDS.items():
        seg = segment_between_keywords(block, pats, allp)
        if not seg:
            continue
        dates = find_dates(seg)
        # Q88 cert table order: Issued, Last Annual, Last Intermediate, Expires. Some rows: Issued, N/A, N/A, Expires.
        if len(dates) >= 4:
            expiry = dates[3]
        elif len(dates) >= 2:
            expiry = dates[-1]
        else:
            continue
        setf(f, key, dstr(expiry), "Q88", "Q88 certificate table row: expiry is last date", seg)

def extract_class_cert_dates(text: str, f: Dict[str, Field]):
    block = cert_block(text)
    # ClassNK table order: cert name, Final, --, Expiry Date, --, Applied.
    class_map = {
        "loadline_expiry": ["Load Line"],
        "safety_construction_expiry": ["Safety Construction"],
        "safety_equipment_expiry": ["Safety Equipment"],
        "safety_radio_expiry": ["Safety Radio"],
        "cof_chem_expiry": ["Chemicals in Bulk"],
        "iopp_expiry": ["OPP (MARPOL Annex I)", "OPP"],
        "ibwmc_expiry": ["BWM"],
    }
    allp = list(class_map.values()) + [["SPP (MARPOL Annex IV)"], ["APP (MARPOL Annex VI)"], ["EE"], ["Anti Fouling"], ["IHM"], ["Lifting Appliances"]]
    for key, pats in class_map.items():
        seg = segment_between_keywords(block, pats, allp)
        dates = find_dates(seg)
        if dates:
            setf(f, key, dstr(dates[0]), "Class Status", "Class status statutory certificate expiry", seg)

# -------------------------
# HVPQ extractor
# -------------------------
def extract_hvpq_xml(xml_text: str) -> Dict[str, Field]:
    f: Dict[str, Field] = {}
    if not xml_text:
        return f
    try:
        root = ET.fromstring(xml_text.encode("utf-8"))
        ns = {"x":"https://www.ocimf-sire.org/schemas/XPQ6_Response_Schema_1.0.00.xsd"}
        v = root.find(".//x:Vessel", ns) or root.find(".//Vessel")
        if v is not None:
            setf(f, "vessel_name", v.attrib.get("name",""), "HVPQ XML", "XML header", ET.tostring(v, encoding="unicode"))
            setf(f, "imo_number", v.attrib.get("id",""), "HVPQ XML", "XML header", ET.tostring(v, encoding="unicode"))
        doc = root.find(".//x:Document", ns) or root.find(".//Document")
        if doc is not None:
            setf(f, "document_exported", doc.attrib.get("exported",""), "HVPQ XML", "XML header", ET.tostring(doc, encoding="unicode"), "MEDIUM")
    except Exception as e:
        setf(f, "xml_parse_error", str(e), "HVPQ XML", "parse error", xml_text[:500], "LOW")
    return f

def first_nonempty_line_after(lines: List[str], token: str) -> str:
    for i, ln in enumerate(lines):
        if token.lower() in ln.lower():
            for j in range(i+1, min(len(lines), i+4)):
                if lines[j].strip(): return lines[j].strip()
    return ""

def extract_hvpq(text: str, xml_text: str="") -> Dict[str, Field]:
    f: Dict[str, Field] = {}
    lines = [clean_value(x) for x in (text or "").splitlines() if clean_value(x)]
    head = text[:1200]
    m = re.search(r"Harmonised Vessel Particulars Questionnaire v6\s+(.*?)\s+IMO/LR Number\s+(\d+)\s+([A-Z0-9 '.-]+)", head, re.S|re.I)
    if m:
        setf(f, "document_date", clean_value(m.group(1)), "HVPQ", "header", m.group(0))
        setf(f, "imo_number", m.group(2), "HVPQ", "header", m.group(0))
        setf(f, "vessel_name", m.group(3), "HVPQ", "header", m.group(0))

    b = extract_qblock(text, "1.1.4", 1200)
    m = re.search(r"\b1\s+Flag\s+([A-Za-z ]+)", b, re.I)
    if m: setf(f,"flag",m.group(1),"HVPQ","qid 1.1.4",b)

    b = extract_qblock(text, "1.1.8", 700)
    m = re.search(r"Q1\.11[^?]*\?\s*([A-Za-z ()/-]+)", b, re.I)
    if m: setf(f,"vessel_type",m.group(1),"HVPQ","qid 1.1.8",b)
    b2 = extract_qblock(text, "1.1.9", 500)
    m = re.search(r"If other,? then specify\s+(.+)", b2, re.I)
    if m: setf(f,"vessel_type_other",m.group(1),"HVPQ","qid 1.1.9",b2)
    if f.get("vessel_type", Field()).value.lower() == "other" and f.get("vessel_type_other"):
        f["vessel_type"] = Field(f["vessel_type_other"].value, "HIGH", "HVPQ", "qid 1.1.8/1.1.9 combined", norm_space(b + "\n" + b2)[:1200])

    b = extract_qblock(text, "1.5.1", 1100)
    m = re.search(r"\b1\s+Classification Society\s+(.+?)\s+\b2\s+Is Classification", b, re.I|re.S)
    if m:
        val = m.group(1)
        val = re.sub(r"^\s*1\s+Classification Society\s+", "", val, flags=re.I).strip()
        setf(f,"class_society",val,"HVPQ","qid 1.5.1",b)

    b = extract_qblock(text, "1.5.2", 1200)
    m = re.search(r"List class notations\s+(.+?)(?:\s+2\s+Provide|\Z)", b, re.I|re.S)
    if m: setf(f,"class_notation",m.group(1),"HVPQ","qid 1.5.2",b)

    # Class and survey dates
    for qid, keys in {
        "1.5.4": ["last_drydock", "second_last_drydock", "next_drydock_due"],
        "1.5.5": ["last_iws", "next_iws_due"],
        "1.5.6": ["last_special_survey", "next_special_survey_due"],
    }.items():
        b = extract_qblock(text, qid, 1400)
        dates = find_dates(b)
        if qid == "1.5.6" and len(dates) >= 2:
            # block also contains society name, next due is last date in the block
            vals = [dates[0], dates[-1]]
            for key, val in zip(keys, vals): setf(f,key,dstr(val),"HVPQ",f"qid {qid}",b)
        else:
            for key, val in zip(keys, dates): setf(f,key,dstr(val),"HVPQ",f"qid {qid}",b)

    b = extract_qblock(text, "1.5.11", 300)
    dates = find_dates(b)
    if dates: setf(f,"last_annual_survey",dstr(dates[0]),"HVPQ","qid 1.5.11",b)
    b = extract_qblock(text, "1.5.12", 300)
    dates = find_dates(b)
    if dates: setf(f,"last_intermediate_survey",dstr(dates[0]),"HVPQ","qid 1.5.12",b)

    for qid, key in [("1.5.14","conditions_of_class"),("1.5.16","memoranda_of_class"),("1.5.18","flag_dispensation")]:
        b = extract_qblock(text, qid, 700)
        yn = answer_yesno(b)
        if yn: setf(f,key,yn,"HVPQ",f"qid {qid}",b)

    # Environmental
    b = extract_qblock(text, "1.2.3", 1300)
    m = re.search(r"provide CII rating\s+([A-Z])\b", b, re.I)
    if m: setf(f,"cii_rating",m.group(1),"HVPQ","qid 1.2.3",b)
    m = re.search(r"CII rating verified by Class, 3rd Party or Owner\?\s*(Class|3rd Party|Owner)", b, re.I)
    if m: setf(f,"cii_verified_by",m.group(1),"HVPQ","qid 1.2.3",b)

    # Incidents and PSC
    b = extract_qblock(text, "1.9.3", 700)
    yn = answer_yesno(b)
    if yn: setf(f,"incident_pollution_grounding_collision",yn,"HVPQ","qid 1.9.3",b)
    b = extract_qblock(text, "1.9.5", 1400)
    yn = answer_yesno(b)
    if yn: setf(f,"incident_other",yn,"HVPQ","qid 1.9.5",b)
    if "Date Type of Incident" in b:
        details = b.split("Date Type of Incident",1)[-1]
        setf(f,"incident_details",details,"HVPQ","qid 1.9.6 table",b,"MEDIUM")

    b = extract_qblock(text, "1.9.8", 1200)
    dates = find_dates(b)
    if dates: setf(f,"last_psc_date",dstr(dates[0]),"HVPQ","qid 1.9.8",b)
    m = re.search(r"Port of last Port State Control Inspection\s+(.+?)\s+3\s+Has", b, re.I|re.S)
    if m: setf(f,"last_psc_port",m.group(1),"HVPQ","qid 1.9.8",b)
    m = re.search(r"detained during the last 36 months\?\s*(Yes|No)", b, re.I)
    if m: setf(f,"psc_detained",m.group(1),"HVPQ","qid 1.9.8",b)

    # Operational checks
    for qid, key, label in [
        ("5.3.1","foam_test_date","foam"), ("6.1.10","overboard_blanks","overboard"), ("6.1.8","sea_chest","sea chest"),
        ("6.1.13","cargo_pressure","cargo pressure"), ("6.1.14","bunker_pressure","bunker pressure")]:
        b = extract_qblock(text, qid, 1300)
        if key == "foam_test_date":
            dates = find_dates(b)
            if dates: setf(f,key,dstr(dates[-1]),"HVPQ",f"qid {qid}",b)
            m = re.search(r"If other, then specify\s+(.+?)\s+4\s+What", b, re.I|re.S)
            if m: setf(f,"foam_type",m.group(1),"HVPQ",f"qid {qid}",b)
        elif key in {"cargo_pressure", "bunker_pressure"}:
            m = re.search(r"specify pressure\s+([0-9.]+\s*Bar)", b, re.I)
            if m: setf(f,key,m.group(1),"HVPQ",f"qid {qid}",b)
        elif key == "overboard_blanks":
            yn = answer_yesno(b)
            if yn: setf(f,key,yn,"HVPQ",f"qid {qid}",b)
        elif key == "sea_chest":
            m = re.search(r"What type of sea valves are fitted.*?\?\s*(.+)", b, re.I|re.S)
            if m: setf(f,key,m.group(1),"HVPQ",f"qid {qid}",b)

    extract_hvpq_cert_dates(text, f)
    for k,v in extract_hvpq_xml(xml_text).items():
        # XML header should improve identity, not override good PDF details unless PDF missing.
        if k not in f or not f[k].value:
            f[k] = v
    return f

# -------------------------
# PIQ extractor
# -------------------------
def extract_piq(text: str) -> Dict[str, Field]:
    f: Dict[str, Field] = {}
    head = text[:900]
    # PIQ header order from PDF extraction is "PIQ Report\nCHALLENGE POLLUX\nVessel Name\nDate\n14 May 2026"
    m = re.search(r"PIQ Report\s+(.+?)\s+Vessel Name\s+Date\s+(.+?)\s+1\. General", head, re.S|re.I)
    if m:
        setf(f,"vessel_name",m.group(1),"PIQ","header",m.group(0))
        setf(f,"document_date",re.split(r"--- PAGE", m.group(2))[0].strip(),"PIQ","header",m.group(0))
    else:
        m = re.search(r"Vessel Name\s+(.+?)\s+Date\s+(.+?)\s+1\. General", head, re.S|re.I)
        if m:
            setf(f,"vessel_name",m.group(1),"PIQ","header",m.group(0))
            setf(f,"document_date",re.split(r"--- PAGE", m.group(2))[0].strip(),"PIQ","header",m.group(0))

    b = extract_qblock(text, "1.1.1", 900)
    m = re.search(r"1\.1\.1\.\s+(.+?)\s+Vessel Type", b, re.S|re.I)
    if not m:
        m = re.search(r"Vessel Type\s+(.+?)(?:\s+If the vessel|\Z)", b, re.S|re.I)
    if m:
        vt = clean_value(m.group(1))
        vt = re.sub(r"\b(Yes|No)\b.*$", "", vt, flags=re.I).strip()
        setf(f,"vessel_type",vt,"PIQ","qid 1.1.1",b)

    b = extract_qblock(text, "2.1.1", 800)
    dates = find_dates(b)
    if dates: setf(f,"class_surveyor_last_visit",dstr(dates[0]),"PIQ","qid 2.1.1",b)
    m = re.search(r"If Other, provide details\s+(.+)", b, re.I|re.S)
    if m: setf(f,"class_surveyor_last_visit_purpose",m.group(1),"PIQ","qid 2.1.1",b)

    # Superintendent visits
    b = extract_qblock(text, "2.2.1001", 1400)
    dates = find_dates(b)
    if dates:
        setf(f,"technical_superintendent_dates", ", ".join(dstr(x) for x in dates), "PIQ", "qid 2.2.1001", b)
        # pairs: from/to for last, second, third
        if len(dates) >= 2: setf(f,"technical_superintendent_last_to", dstr(dates[1]), "PIQ", "qid 2.2.1001", b)
    b = extract_qblock(text, "2.2.1002", 1000)
    dates = find_dates(b)
    if dates:
        setf(f,"marine_superintendent_dates", ", ".join(dstr(x) for x in dates), "PIQ", "qid 2.2.1002", b)
        if len(dates) >= 2: setf(f,"marine_superintendent_last_to", dstr(dates[1]), "PIQ", "qid 2.2.1002", b)

    for qid,key in [("2.3.3001","cargo_tank_oldest_inspection"),("2.3.3002","ballast_tank_oldest_inspection"),("2.3.3003","void_oldest_inspection")]:
        b = extract_qblock(text, qid, 700)
        dates = find_dates(b)
        if dates: setf(f,key,dstr(dates[-1]),"PIQ",f"qid {qid}",b)
        m = re.search(r"Required frequency.*?\?\s*(\d+)\s+months", b, re.I|re.S)
        if m: setf(f,key+"_frequency_months",m.group(1),"PIQ",f"qid {qid}",b)

    b = extract_qblock(text,"2.5.1002",1000)
    yn = answer_yesno(b)
    if yn: setf(f,"equipment_retrofitted",yn,"PIQ","qid 2.5.1002",b)
    if "EGCS" in b.upper(): setf(f,"egcs_retrofit_mentioned","Yes","PIQ","qid 2.5.1002",b)

    # PSC table: take Last row only: Last 03 January 2026 Chattogram , Bangladesh Indian Ocean MoU 0.00 No Yes
    b = extract_qblock(text, "2.8.2", 1800)
    m = re.search(r"Last\s+({d})\s+(.+?)\s+(?:Indian Ocean MoU|Tokyo MoU|US Coastguard|Paris MoU|USCG)\s+([0-9.]+)\s+(Yes|No)\s+(Yes|No)".format(d=DATE_RE.pattern), b, re.I|re.S)
    if m:
        dt = find_dates(m.group(1))[0]
        setf(f,"last_psc_date",dstr(dt),"PIQ","qid 2.8.2 last row",b)
        port = clean_value(re.sub(r",.*", "", m.group(2)))
        setf(f,"last_psc_port",port,"PIQ","qid 2.8.2 last row",b)
        setf(f,"psc_detained",m.group(4),"PIQ","qid 2.8.2 last row",b)
        setf(f,"last_psc_deficiencies",m.group(3),"PIQ","qid 2.8.2 last row",b)

    # Incident section: PIQ 5.7.*. If all specified incidents are No, flag nil, but compare broad HVPQ other incident separately.
    sec = re.search(r"5\.7\. Safety Management(.*?)(?:\n\s*5\.8\.|\n\s*6\.|\Z)", text, re.S|re.I)
    if sec:
        block = sec.group(1)
        setf(f,"piq_incident_section",block,"PIQ","section 5.7",block,"MEDIUM")
        positives = re.findall(r"5\.7\.\d+\.\s+(Yes)", block, re.I)
        if positives:
            setf(f,"incident_other","Yes","PIQ","section 5.7 any Yes",block)
        else:
            setf(f,"incident_other","No","PIQ","section 5.7 all extracted incident answers appear No",block)
        # specific pollution/grounding/collision items in 5.7 usually separate; broad group No if all these keywords No.
        if re.search(r"pollution incident.*?No", block, re.S|re.I) and re.search(r"hard aground.*?No", block, re.S|re.I) and re.search(r"collision or allision.*?No", block, re.S|re.I):
            setf(f,"incident_pollution_grounding_collision","No","PIQ","section 5.7 specific incident answers",block)
    return f

# -------------------------
# Q88 extractor
# -------------------------
def q88_block(text: str, qnum: str, max_chars=1200) -> str:
    """Return Q88 field block by line-based question number matching.
    Handles Q88 PDF extraction where qid is often on its own line and the label/value follow on next lines.
    """
    lines = [clean_value(x) for x in (text or "").splitlines()]
    start_idx = None
    q_re = re.compile(rf"^{re.escape(qnum)}(?:\b|\s*$)", re.I)
    next_re = re.compile(r"^\d+\.\d+[a-z]?(?:\b|\s*$)", re.I)
    for i, ln in enumerate(lines):
        if q_re.search(ln):
            start_idx = i
            break
    if start_idx is None:
        return ""
    out=[]
    for j in range(start_idx, min(len(lines), start_idx+80)):
        if j > start_idx and next_re.search(lines[j]):
            break
        out.append(lines[j])
    block = " ".join([x for x in out if x])
    return block[:max_chars]

def extract_q88(text: str) -> Dict[str, Field]:
    f: Dict[str, Field] = {}
    if not text: return f
    # Header/general fields by explicit question numbers.
    qmap = {
        "vessel_name": ("1.2", r"Vessel.*?\(IMO number\)\s*(.+?)\s*\((\d+)\)"),
        "flag": ("1.5", r"Flag/Port of Registry\s*(.+?)/(.*)"),
        "vessel_type": ("1.8", r"IOPPC\)\s*(.+)"),
        "class_society": ("1.18", r"Classification society\s*(.+?)(?:\s+1\.18a|\Z)"),
        "class_notation": ("1.19", r"Class notation\s*(.+?)(?:\s+1\.20|\Z)"),
        "conditions_of_class": ("1.20", r"open conditions.*?\b(Yes|No)\s*(?:1\.20a|$)"),
        "memoranda_of_class": ("1.20a", r"Memoranda of Class.*?\b(Yes|No)\s*(?:1\.21|$)"),
    }
    for key,(qid,pat) in qmap.items():
        b = q88_block(text,qid,1300)
        if not b: continue
        m = re.search(pat,b,re.S|re.I)
        if m:
            if key == "vessel_name":
                setf(f,"vessel_name",m.group(1),"Q88",f"q88 {qid}",b)
                setf(f,"imo_number",m.group(2),"Q88",f"q88 {qid}",b)
            elif key == "flag":
                setf(f,"flag",m.group(1),"Q88",f"q88 {qid}",b)
            elif key in {"conditions_of_class", "memoranda_of_class"}:
                yns = re.findall(r"\b(Yes|No)\b", b, flags=re.I)
                setf(f,key,yns[-1] if yns else m.group(1),"Q88",f"q88 {qid}",b)
            else:
                val = m.group(1)
                if key == "class_society": val = re.split(r"\s+1\.18a", val)[0]
                setf(f,key,val,"Q88",f"q88 {qid}",b)

    for qid, keys in {"1.23":["last_drydock"], "1.24":["next_drydock_due","next_annual_survey_due"], "1.25":["last_special_survey","next_special_survey_due"], "1.25a":["last_iws","next_iws_due"]}.items():
        b = q88_block(text,qid,900)
        dates = find_dates(b)
        if qid == "1.23" and dates:
            setf(f,"last_drydock",dstr(dates[0]),"Q88",f"q88 {qid}",b)
        elif qid == "1.24" and dates:
            setf(f,"next_drydock_due",dstr(dates[0]),"Q88",f"q88 {qid}",b)
            if len(dates)>1: setf(f,"next_annual_survey_due",dstr(dates[1]),"Q88",f"q88 {qid}",b)
        elif len(dates)>=2:
            setf(f,keys[0],dstr(dates[0]),"Q88",f"q88 {qid}",b)
            setf(f,keys[1],dstr(dates[1]),"Q88",f"q88 {qid}",b)

    # P&I and foam/cargo fields
    b = q88_block(text,"1.14",1000)
    if b:
        m=re.search(r"P\s*&\s*I Club.*?:\s*(.+)",b,re.I|re.S)
        if m: setf(f,"pni_club",m.group(1),"Q88","q88 1.14",b)
    b = q88_block(text,"1.15",500)
    dates=find_dates(b)
    if dates: setf(f,"pni_expiry",dstr(dates[-1]),"Q88","q88 1.15",b)

    # Q88 operational pages often use numbered rows; generic label extraction.
    for label,key in [("foam", "foam_type"),("test", "foam_test_date"),("overboard", "overboard_blanks"),("sea chest", "sea_chest"), ("cargo piping", "cargo_pressure"), ("bunker piping", "bunker_pressure")]:
        m = re.search(label + r".{0,200}", text, re.I|re.S)
        if m:
            seg = text[m.start():m.start()+600]
            if key.endswith("date"):
                dates=find_dates(seg)
                if dates: setf(f,key,dstr(dates[0]),"Q88","keyword extraction",seg,"MEDIUM")
            elif key.endswith("pressure"):
                mm=re.search(r"([0-9.]+\s*Bar)",seg,re.I)
                if mm: setf(f,key,mm.group(1),"Q88","keyword extraction",seg,"MEDIUM")
            else:
                # don't set low-quality broad fields unless answer is obvious yes/no/type
                yn=re.search(r"\b(Yes|No|AR-AFFF\s*3\s*%|Screwdown)\b",seg,re.I)
                if yn: setf(f,key,yn.group(1),"Q88","keyword extraction",seg,"MEDIUM")
    extract_q88_cert_dates(text,f)
    b = q88_block(text, "2.19", 600)
    if b:
        ds = find_dates(b)
        if ds: f["cof_chem_expiry"] = Field(dstr(ds[-1]), "HIGH", "Q88", "q88 2.19 Certificate of Fitness", norm_space(b)[:1200])
    b = q88_block(text, "2.11", 500)
    if b:
        ds = find_dates(b)
        if ds: f["uscg_coc_expiry"] = Field(dstr(ds[-1]), "HIGH", "Q88", "q88 2.11 USCGCOC", norm_space(b)[:1200])
    return f

# -------------------------
# Class status extractor
# -------------------------
def extract_class_status(text: str) -> Dict[str, Field]:
    f: Dict[str, Field] = {}
    if not text: return f
    head = text[:2500]
    m = re.search(r"Name of Ship:\s*\n?\s*(.+?)\s+(?:Class No|IMO No)", head, re.I|re.S)
    if m: setf(f,"vessel_name",m.group(1),"Class Status","header",m.group(0))
    m = re.search(r"IMO No\.\s*:?\s*(\d{7})", head, re.I)
    if m: setf(f,"imo_number",m.group(1),"Class Status","header",m.group(0))
    if re.search(r"NIPPON KAIJI KYOKAI|ClassNK|NK-SHIPS", text, re.I):
        setf(f,"class_society","Nippon Kaiji Kyokai","Class Status","document source",head)
    m = re.search(r"Flag:\s*([^\n]+)", text, re.I)
    if m: setf(f,"flag",m.group(1),"Class Status","particulars",m.group(0))
    m = re.search(r"Type of Ship -Purpose\(Intended Service\):\s*\n?\s*([^\n]+)", text, re.I)
    if m:
        val = m.group(1)
        if val.strip().startswith("-") or "Certificates" in val:
            mm = re.search(r"SC/SE/SF:\s*([^\n]+)", text, re.I)
            if mm: val = mm.group(1)
        setf(f,"vessel_type",val,"Class Status","particulars",m.group(0))
    m = re.search(r"Classification Character, Notations:\s*(.+?)\s+Descriptive Notes", text, re.I|re.S)
    if m:
        val = m.group(1)
        # If layout puts the heading before the actual notation, also grab the NS*/MNS* lines above the heading.
        mm = re.search(r"(NS\*.*?MNS\*)\s*Classification Character, Notations", text, re.I|re.S)
        if mm: val = mm.group(1)
        setf(f,"class_notation",val,"Class Status","particulars",m.group(0))

    # Survey status page. Dates extracted per row.
    def row(name: str, max_len=500) -> str:
        mm = re.search(re.escape(name) + r"(.{0,"+str(max_len)+r"})", text, re.I|re.S)
        return name + (mm.group(1) if mm else "")
    for name,last_key,due_key in [
        ("Special Survey","last_special_survey","next_special_survey_due"),
        ("Annual Survey","last_annual_survey",None),
        ("Intermediate Survey","last_intermediate_survey",None),
        ("Docking Survey","last_iws","next_iws_due"),
    ]:
        seg=row(name,450)
        dates=find_dates(seg)
        # ClassNK Survey Status order is: Last Date, Due Date, Range/Postponed.
        if dates:
            setf(f,last_key,dstr(dates[0]),"Class Status",f"Survey Status row: {name}",seg)
            if due_key and len(dates)>=2:
                setf(f,due_key,dstr(dates[1]),"Class Status",f"Survey Status row: {name}",seg)
    # Drydock in HVPQ generally equivalent to Docking Survey in Class Status
    if f.get("last_iws"):
        setf(f,"last_drydock",f["last_iws"].value,"Class Status","Docking Survey used as class docking survey last date",f["last_iws"].evidence,"MEDIUM")
    if f.get("next_iws_due"):
        setf(f,"next_drydock_due",f["next_iws_due"].value,"Class Status","Docking Survey used as class docking survey due date",f["next_iws_due"].evidence,"MEDIUM")

    # Conditions and notes
    m = re.search(r"Condition of Class\s*(Nil\.|None|No|.+?)\s*Note", text, re.I|re.S)
    if m:
        v=m.group(1)
        setf(f,"conditions_of_class", "No" if re.search(r"Nil|None|No",v,re.I) else v, "Class Status", "Condition of Class", m.group(0))
    m = re.search(r"\bNote\s*(Nil\.|None|No|.+?)\s*2\. Installation", text, re.I|re.S)
    if m:
        v=m.group(1)
        setf(f,"memoranda_of_class", "No" if re.search(r"Nil|None|No",v,re.I) else v, "Class Status", "Class Note", m.group(0))
    extract_class_cert_dates(text,f)
    return f

# -------------------------
# Observation library and checklist
# -------------------------
def parse_obs_excel(uploaded) -> pd.DataFrame:
    if uploaded is None: return pd.DataFrame()
    try:
        df = pd.read_excel(uploaded)
    except Exception:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    if df.empty: return df
    # Robust row text join. pandas versions differ in how DataFrame.agg handles string-callables.
    # Using apply with explicit conversion avoids TypeError on Streamlit Cloud / pandas 3.x.
    joined = df.fillna("").astype(str).apply(lambda row: " ".join([str(v) for v in row.tolist()]), axis=1)
    qids = joined.str.extract(r"((?:\d+\.){1,4}\d+)")[0]
    df["detected_qid"] = qids
    def fam(s):
        x=s.lower()
        if any(w in x for w in ["incident", "ground", "blackout", "injury", "collision", "allision"]): return "Incident declaration"
        if any(w in x for w in ["certificate", "expiry", "issued", "annual", "intermediate"]): return "Certificate/date accuracy"
        if any(w in x for w in ["class", "condition", "memoranda", "dispensation", "dry dock", "iws"]): return "Class/survey"
        if any(w in x for w in ["mooring", "brake", "rope", "tail", "winch", "bitt", "fairlead", "chock"]): return "Mooring"
        if any(w in x for w in ["tank", "coating", "void", "ballast"]): return "Tank/structural"
        if any(w in x for w in ["foam", "fire", "lifeboat", "rescue"]): return "Firefighting/LSA"
        if any(w in x for w in ["piping", "overboard", "sea chest", "scupper", "pressure"]): return "Pollution/cargo"
        if any(w in x for w in ["diagram", "manifold", "arrangement"]): return "Diagrams"
        return "Other HVPQ accuracy"
    df["family"] = joined.map(fam)
    return df

def targeted_checklist(obs: pd.DataFrame) -> pd.DataFrame:
    base = [
        ("PIQ/HVPQ", "Incident declaration", "Confirm HVPQ 1.9.3/1.9.5 and PIQ 5.7 answers reflect all incidents in previous 12 months. If nil, vessel must positively confirm nil."),
        ("PIQ/HVPQ", "PSC", "Confirm last 3 PSC inspections in PIQ and last PSC in HVPQ are updated: date, port, MOU, deficiencies, detention, OCIMF entry."),
        ("Class Status", "Class/certificates", "Verify HVPQ/Q88 certificate expiry dates against latest Class Status and certificates."),
        ("Class Status", "COC/MOC/dispensation", "Verify Conditions of Class, Memoranda/Notes of Class and flag/class dispensations are correctly declared."),
        ("PIQ", "Superintendent visits", "Verify Technical Superintendent gap <=7 months and Marine Superintendent gap <=12 months; any exceedance must be explained/actioned."),
        ("PIQ", "Tank inspections", "Verify cargo/slop, ballast and void inspection sequence dates and due dates against tank inspection records."),
        ("HVPQ/Q88", "Mooring", "Verify mooring winch details, brake test date, brake holding capacity/rendering load, split drum, rope/tail certificates and end-for-end/discard records."),
        ("HVPQ/Q88", "Cargo/pollution", "Verify cargo/bunker pressure test, overboard blanks/testing, sea chest, scupper plugs and cargo system details."),
        ("HVPQ/Q88", "Firefighting", "Verify foam type/test certificate date, fixed systems, sample locker systems and rescue boat/davit declarations."),
        ("HVPQ/Q88", "Diagrams", "Verify mooring layout, manifold layout, fairlead/chock/bitt and bow mooring arrangement diagrams are attached/current."),
    ]
    rows = [{"source":s,"area":a,"ship_check":c,"priority":"Targeted"} for s,a,c in base]
    if obs is not None and not obs.empty and "family" in obs:
        counts=obs["family"].value_counts().to_dict()
        for fam,cnt in counts.items():
            rows.append({"source":"Observation library","area":fam,"ship_check":f"Historical observation family appears {cnt} time(s). Include this area in targeted ship verification.","priority":"Observation-driven"})
    return pd.DataFrame(rows)

# -------------------------
# Rules/comparison
# -------------------------
def compare_pair(field: str, label: str, area: str, a_name: str, a: Field, b_name: str, b: Field) -> Optional[Finding]:
    if not a.value or not b.value:
        return None
    ok, reason = equivalent(a.value, b.value, field)
    if ok:
        return None
    risk = "CRITICAL" if field in {"imo_number", "psc_detained", "incident_pollution_grounding_collision"} else "HIGH"
    values = {"hvpq_value":"", "piq_value":"", "class_value":"", "q88_value":""}
    for nm, fld in [(a_name,a), (b_name,b)]:
        key = {"hvpq":"hvpq_value", "piq":"piq_value", "class":"class_value", "q88":"q88_value"}.get(nm, "hvpq_value")
        values[key] = fld.value
    return Finding(area, f"{label}: {a_name.upper()} vs {b_name.upper()}", "MISMATCH", risk,
                   **values, reason=reason,
                   required_action="Verify against source evidence and correct the declaration if needed.",
                   evidence=f"{a_name.upper()}: {a.evidence}\n{b_name.upper()}: {b.evidence}")

def add_missing_source_checks(hvpq, piq, cls, q88) -> List[Finding]:
    out=[]
    # CII verified by Owner is a warning/manual check, not a mismatch.
    if hvpq.get("cii_verified_by", Field()).value.lower() == "owner":
        out.append(Finding("Environmental", "CII verification basis", "MANUAL CHECK", "MEDIUM",
            hvpq_value=f"CII {hvpq.get('cii_rating', Field()).value}; verified by Owner",
            reason="CII value is declared but verification basis is Owner. Inspectors often expect supporting latest CII/SEEMP evidence.",
            required_action="Vessel/office to verify latest CII/AER evidence and confirm HVPQ entry is current.",
            evidence=hvpq.get("cii_verified_by", Field()).evidence))
    return out

def run_comparison(hvpq: Dict[str, Field], piq: Dict[str, Field], cls: Dict[str, Field], q88: Dict[str, Field]) -> List[Finding]:
    srcs={"hvpq":hvpq,"piq":piq,"class":cls,"q88":q88}
    findings=[]
    for field, area, label, allowed_sources in COMPARE_MAP:
        # Only compare HVPQ against other docs, except PIQ-only checks below. Do not compare Q88 vs Class etc.
        if "hvpq" in allowed_sources and hvpq.get(field, Field()).value:
            for other in allowed_sources:
                if other == "hvpq": continue
                fnd = compare_pair(field,label,area,"hvpq",hvpq[field],other,srcs[other].get(field, Field()))
                if fnd: findings.append(fnd)
    # Certificates: HVPQ vs Class/Q88 only when both extracted confidently.
    for key in CERT_COMPARE_KEYS:
        hv = hvpq.get(key, Field())
        if not hv.value: continue
        label = key.replace("_", " ").replace("expiry", "expiry").title()
        for other, odict in [("class",cls),("q88",q88)]:
            if key in odict and odict[key].value:
                fnd = compare_pair(key,label,"Certificates","hvpq",hv,other,odict[key])
                if fnd: findings.append(fnd)
    # Incident nil/declared mismatch HVPQ vs PIQ specific
    h_inc = hvpq.get("incident_other", Field()).value
    p_inc = piq.get("incident_other", Field()).value
    if h_inc and p_inc and norm_for_compare(h_inc) != norm_for_compare(p_inc):
        findings.append(Finding("Incidents","Other incident declaration: HVPQ vs PIQ","MISMATCH","HIGH",
            hvpq_value=h_inc, piq_value=p_inc,
            reason="HVPQ and PIQ do not appear aligned on previous-12-month incident declaration.",
            required_action="Ship/office to confirm incident history and update both HVPQ and PIQ consistently.",
            evidence=f"HVPQ: {hvpq.get('incident_other').evidence}\nPIQ: {piq.get('piq_incident_section', Field()).evidence}"))
    if h_inc.lower()=="no" and (not p_inc or p_inc.lower()=="no"):
        findings.append(Finding("Incidents","No incidents declared","MANUAL CHECK","MEDIUM",
            hvpq_value=h_inc, piq_value=p_inc,
            reason="No incidents appear declared. Positive confirmation is required because incident non-reporting is a recurring observation category.",
            required_action="Ask vessel/office to confirm no reportable incidents in previous 12 months before submission."))

    # Superintendent rule: deterministic PIQ-only
    for key, label, max_months in [("technical_superintendent_last_to","Technical Superintendent inspection gap",7),("marine_superintendent_last_to","Marine Superintendent inspection gap",12)]:
        val=piq.get(key, Field()).value
        docdate=piq.get("document_date", Field()).value or hvpq.get("document_date", Field()).value
        if val:
            try:
                last=dateparser.parse(val).date(); today=dateparser.parse(docdate, dayfirst=True).date() if docdate else date.today()
                due=last+relativedelta(months=max_months)
                if today>due:
                    findings.append(Finding("Management Oversight",label,"MISMATCH","CRITICAL",piq_value=f"Last to: {val}; due by: {dstr(due)}; document date: {dstr(today)}",reason=f"Gap exceeds strict {max_months}.0 month rule.",required_action="Arrange/justify superintendent visit and correct PIQ if needed.",evidence=piq.get(key).evidence))
            except Exception:
                pass
    # Tank inspection due calculations PIQ-only, always manual/targeted
    for dk,label in [("cargo_tank_oldest_inspection","Cargo/slop tank"),("ballast_tank_oldest_inspection","Ballast tank"),("void_oldest_inspection","Void space")]:
        if dk in piq:
            freq=int(float(piq.get(dk+"_frequency_months", Field("12")).value or 12))
            try:
                d=dateparser.parse(piq[dk].value).date(); due=d+relativedelta(months=freq)
                findings.append(Finding("Tank Inspection",f"{label} inspection sequence due calculation","MANUAL CHECK","MEDIUM",piq_value=f"Oldest date: {dstr(d)}; frequency: {freq} months; next due: {dstr(due)}",reason="PIQ tank inspection sequence should be verified against latest tank inspection records.",required_action="Ship to confirm latest tank inspection records and update PIQ if sequence changed.",evidence=piq[dk].evidence))
            except Exception: pass
    findings.extend(add_missing_source_checks(hvpq,piq,cls,q88))
    # Deduplicate exact check/source/value combos
    seen=set(); uniq=[]
    for f in findings:
        key=(f.area,f.check,f.hvpq_value,f.piq_value,f.class_value,f.q88_value)
        if key not in seen:
            seen.add(key); uniq.append(f)
    return uniq

# -------------------------
# UI and exports
# -------------------------
def fields_df(name: str, d: Dict[str, Field]) -> pd.DataFrame:
    return pd.DataFrame([{ "source_doc": name, "field": k, **asdict(v)} for k,v in sorted(d.items())])

def make_excel(findings: List[Finding], checklist: pd.DataFrame, extracted: pd.DataFrame, obs: pd.DataFrame) -> bytes:
    bio=io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        pd.DataFrame([asdict(x) for x in findings]).to_excel(writer,index=False,sheet_name="Findings")
        checklist.to_excel(writer,index=False,sheet_name="Ship Checklist")
        extracted.to_excel(writer,index=False,sheet_name="Extracted Fields")
        if obs is not None and not obs.empty:
            obs.to_excel(writer,index=False,sheet_name="Observation Library")
    return bio.getvalue()

def main():
    st.set_page_config(page_title=f"HVPQ / PIQ Checker {APP_VERSION}", layout="wide")
    st.title(f"HVPQ / PIQ Vetting Observation Checker {APP_VERSION}")
    st.caption("Extraction-first. HVPQ/Q88 mapped by section. Class Status limited to certificates, survey status, COC/MOC/notes. No mismatch if extraction is missing.")
    with st.sidebar:
        hvpq_file=st.file_uploader("HVPQ PDF", type=["pdf"])
        piq_file=st.file_uploader("PIQ PDF", type=["pdf"])
        class_file=st.file_uploader("Class Status PDF", type=["pdf"])
        q88_file=st.file_uploader("Q88 PDF", type=["pdf"])
        xml_file=st.file_uploader("HVPQ XML optional", type=["xml"])
        obs_file=st.file_uploader("HVPQ observation Excel optional", type=["xlsx"])
        inc_obs_file=st.file_uploader("Incident observation Excel optional", type=["xlsx"])
        show_extraction_debug=st.checkbox("Show raw text debug", False)
        run=st.button("Run extraction and checks", type="primary")
    if not run:
        st.info("Upload documents and run checks.")
        return
    hvpq_text=read_pdf(hvpq_file) if hvpq_file else ""
    piq_text=read_pdf(piq_file) if piq_file else ""
    class_text=read_pdf(class_file) if class_file else ""
    q88_text=read_pdf(q88_file) if q88_file else ""
    xml_text=xml_file.getvalue().decode("utf-8", errors="ignore") if xml_file else ""
    hvpq=extract_hvpq(hvpq_text, xml_text)
    piq=extract_piq(piq_text)
    cls=extract_class_status(class_text)
    q88=extract_q88(q88_text)
    obs=pd.concat([parse_obs_excel(obs_file), parse_obs_excel(inc_obs_file)], ignore_index=True) if (obs_file or inc_obs_file) else pd.DataFrame()
    checklist=targeted_checklist(obs)
    findings=run_comparison(hvpq,piq,cls,q88)
    df=pd.DataFrame([asdict(x) for x in findings])
    extracted=pd.concat([fields_df("HVPQ",hvpq), fields_df("PIQ",piq), fields_df("Class Status",cls), fields_df("Q88",q88)], ignore_index=True)

    st.subheader("Actionable mismatch / manual-check register")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Actionable rows", len(df))
    c2.metric("Critical", int((df['risk'].eq('CRITICAL')).sum()) if not df.empty else 0)
    c3.metric("High", int((df['risk'].eq('HIGH')).sum()) if not df.empty else 0)
    c4.metric("Manual/Medium", int((df['risk'].isin(['MEDIUM','MANUAL CHECK'])).sum()) if not df.empty else 0)
    tabs=st.tabs(["Findings", "Extracted fields", "Ship checklist", "Observation library", "Debug"])
    with tabs[0]:
        if df.empty: st.success("No actionable mismatch found from extracted fields. Review extracted fields/checklist before sending to vessel.")
        else: st.dataframe(df, use_container_width=True, height=540)
    with tabs[1]:
        st.dataframe(extracted, use_container_width=True, height=620)
    with tabs[2]:
        st.dataframe(checklist, use_container_width=True, height=420)
    with tabs[3]:
        if obs.empty: st.caption("No observation library uploaded.")
        else: st.dataframe(obs, use_container_width=True, height=500)
    with tabs[4]:
        if show_extraction_debug:
            st.text_area("HVPQ text preview", hvpq_text[:5000], height=200)
            st.text_area("PIQ text preview", piq_text[:5000], height=200)
            st.text_area("Class Status text preview", class_text[:5000], height=200)
            st.text_area("Q88 text preview", q88_text[:5000], height=200)
        else:
            st.caption("Enable 'Show raw text debug' in sidebar to inspect extraction text.")
    st.download_button("Download Excel register", data=make_excel(findings, checklist, extracted, obs), file_name="hvpq_piq_vetting_register.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
