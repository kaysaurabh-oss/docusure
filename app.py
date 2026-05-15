from __future__ import annotations

import io
import json
import re
import zipfile
from dataclasses import dataclass, asdict
from datetime import date, datetime
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

import pandas as pd
import pymupdf as fitz
import requests
import streamlit as st
from dateutil import parser as dateparser
from dateutil.relativedelta import relativedelta

APP_TITLE = "HVPQ / PIQ / Q88 / Class Status Checker v9"
APP_SUBTITLE = "Extraction-first, optional local-LLM assisted, rule-based verification register."

# ----------------------------- Data models -----------------------------

@dataclass
class FieldRecord:
    source: str
    field_id: str
    label: str
    value: str = ""
    date_value: str = ""
    confidence: str = "deterministic"
    raw: str = ""

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
    action: str = ""

# ----------------------------- General helpers -----------------------------

MONTHS_7_DAYS = 7 * 30.4375
MONTHS_12_DAYS = 12 * 30.4375
DATE_PATTERNS = [
    r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b",
    r"\b[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}\b",
    r"\b\d{4}-\d{2}-\d{2}\b",
    r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b",
]
DATE_RE = re.compile("|".join(DATE_PATTERNS), re.I)

CERT_ALIASES = {
    "safety equipment certificate": "safety_equipment",
    "sec": "safety_equipment",
    "safety radio certificate": "safety_radio",
    "src": "safety_radio",
    "safety construction certificate": "safety_construction",
    "scc": "safety_construction",
    "loadline certificate": "loadline",
    "load line": "loadline",
    "international loadline certificate": "loadline",
    "international oil pollution prevention": "iopp",
    "oil pollution prevention": "iopp",
    "ioppc": "iopp",
    "opp": "iopp",
    "international ballast water management": "bwm",
    "ballast water management": "bwm",
    "ibwmc": "bwm",
    "safety management certificate": "smc",
    "ism safety management certificate": "smc",
    "smc": "smc",
    "document of compliance": "doc",
    "doc": "doc",
    "international ship security": "issc",
    "issc": "issc",
    "uscg certificate of compliance": "uscg_coc",
    "us certificate of compliance": "uscg_coc",
    "uscgcoc": "uscg_coc",
    "civil liability convention certificate": "clc_oil",
    "civil liability certificates": "clc_oil",
    "civil liability convention 1992": "clc_oil",
    "bunker oil pollution": "clc_bunker",
    "bunker": "clc_bunker",
    "wreck removal": "wreck_removal",
    "wrc": "wreck_removal",
    "financial responsibility": "cofr",
    "cofr": "cofr",
    "vessel general permit": "vgp",
    "certificate of fitness": "cof_chemical",
    "cof": "cof_chemical",
    "certificate of class": "class_certificate",
    "class certificate": "class_certificate",
    "international tonnage": "tonnage",
    "international air pollution prevention": "iapp",
    "iapp": "iapp",
    "air pollution prevention": "iapp",
    "international sewage pollution prevention": "ispp",
    "ispp": "ispp",
    "maritime labour certificate": "mlc",
    "mlc": "mlc",
    "ship sanitation": "ship_sanitation",
}

FIELD_LABELS = {
    "vessel.name": "Vessel name",
    "vessel.imo": "IMO number",
    "vessel.flag": "Flag",
    "vessel.port_registry": "Port of registry",
    "vessel.type": "Vessel type",
    "vessel.call_sign": "Call sign",
    "vessel.mmsi": "MMSI",
    "owner.registered_owner": "Registered owner",
    "owner.technical_operator": "Technical operator",
    "owner.commercial_operator": "Commercial operator",
    "insurance.pni_club": "P&I club",
    "classification.class_society": "Class society",
    "classification.class_notation": "Class notation",
    "classification.conditions_of_class": "Conditions of class",
    "classification.memo_of_class": "Memoranda of class",
    "classification.flag_dispensation": "Flag/Class dispensation",
    "surveys.last_drydock": "Last dry dock",
    "surveys.next_drydock_due": "Next dry dock due",
    "surveys.last_iws": "Last IWS",
    "surveys.next_iws_due": "Next IWS due",
    "surveys.last_special": "Last special survey",
    "surveys.next_special_due": "Next special survey due",
    "surveys.last_annual": "Last annual survey",
    "surveys.next_annual_due": "Next annual survey due",
    "surveys.last_intermediate": "Last intermediate survey",
    "environment.cii_rating": "CII rating",
    "environment.cii_verified_by": "CII verification basis",
    "environment.eexi_rating": "EEXI rating",
    "environment.eexi_verified_by": "EEXI verification basis",
    "incidents.pollution_grounding_collision_allision": "Pollution/grounding/collision/allision incident",
    "incidents.other_incidents": "Other incidents",
    "psc.last_date": "Last PSC date",
    "psc.last_port": "Last PSC port",
    "psc.detained_36m": "Detained last 36 months",
    "moc.retrofit": "Equipment retrofitted",
    "moc.structural_change": "Structural change",
    "moc.equipment_replaced": "Equipment replaced non-like-for-like",
}


def clean_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).replace("\uFFFE", " ").replace("\u0000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_key(s: str) -> str:
    s = clean_text(s).lower()
    s = s.replace("&", "and")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_key(a), normalize_key(b)).ratio()


def normalize_value(s: str) -> str:
    s = clean_text(s).lower()
    s = s.replace("limited", "ltd").replace("limitied", "ltd").replace("co.,", "co")
    s = s.replace("n/a", "na").replace("not applicable", "na")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def normalize_bool(s: str) -> str:
    x = normalize_value(s)
    if x in ["yes", "y", "true", "1"]:
        return "yes"
    if x in ["no", "n", "false", "0", "nil", "none"]:
        return "no"
    if x in ["na", "not applicable", "blank", ""]:
        return ""
    return x


def parse_date_any(s: str) -> Optional[date]:
    if not s:
        return None
    # Avoid question numbers / decimals being parsed as dates
    s = clean_text(s)
    m = DATE_RE.search(s)
    target = m.group(0) if m else s
    try:
        dt = dateparser.parse(target, dayfirst=False, fuzzy=True)
        if dt and 1900 <= dt.year <= 2100:
            return dt.date()
    except Exception:
        pass
    return None


def iso_date(s: str) -> str:
    d = parse_date_any(s)
    return d.isoformat() if d else ""


def extract_dates(line: str) -> List[str]:
    out = []
    for m in DATE_RE.finditer(line or ""):
        d = parse_date_any(m.group(0))
        if d:
            out.append(d.isoformat())
    return out


def add_field(fields: List[FieldRecord], source: str, field_id: str, value: Any, label: str = "", raw: str = "", confidence: str = "deterministic"):
    val = clean_text(value)
    if val == "":
        return
    fields.append(FieldRecord(source=source, field_id=field_id, label=label or FIELD_LABELS.get(field_id, field_id), value=val, date_value=iso_date(val), raw=clean_text(raw)[:800], confidence=confidence))


def first_field(fields: List[FieldRecord], source: str, field_id: str) -> str:
    for f in fields:
        if f.source == source and f.field_id == field_id and clean_text(f.value):
            return f.value
    return ""


def values_by_field(fields: List[FieldRecord], source: str) -> Dict[str, str]:
    d = {}
    for f in fields:
        if f.source == source and f.field_id not in d and clean_text(f.value):
            d[f.field_id] = f.value
    return d


def sources_value(fields: List[FieldRecord], field_id: str) -> Dict[str, str]:
    out = {}
    for f in fields:
        if f.field_id == field_id and clean_text(f.value):
            out.setdefault(f.source, f.value)
    return out


def cert_key_from_label(label: str) -> Optional[str]:
    nk = normalize_key(label)
    if "garbage pollution" in nk:
        return None
    # longer aliases first to avoid bunker matching everything
    for alias, key in sorted(CERT_ALIASES.items(), key=lambda kv: len(kv[0]), reverse=True):
        if normalize_key(alias) in nk:
            return key
    return None


def semantically_equivalent(a: str, b: str, field_id: str = "") -> bool:
    if not clean_text(a) or not clean_text(b):
        return False
    da, db = parse_date_any(a), parse_date_any(b)
    if da and db:
        return da == db
    na, nb = normalize_bool(a), normalize_bool(b)
    if na and nb and na == nb:
        return True
    va, vb = normalize_value(a), normalize_value(b)
    if not va or not vb:
        return False
    if va == vb:
        return True
    # Class aliases
    if field_id.endswith("class_society"):
        class_aliases = [
            ("korean register", "kr"), ("nippon kaiji kyokai", "class nk"), ("nippon kaiji kyokai", "nk"),
            ("dnv gl", "dnv"), ("american bureau of shipping", "abs"), ("bureau veritas", "bv"),
        ]
        for x, y in class_aliases:
            if (x in va and y in vb) or (x in vb and y in va):
                return True
    # Vessel type should be broad equivalence, not hard mismatch
    if field_id.endswith("vessel.type"):
        oil_words_a = any(w in va for w in ["oil", "crude", "product", "chemical", "tanker", "carrier"])
        oil_words_b = any(w in vb for w in ["oil", "crude", "product", "chemical", "tanker", "carrier"])
        if oil_words_a and oil_words_b:
            return True
    # P&I / owner fuzzy equivalence
    if similarity(va, vb) >= 0.88:
        return True
    return False

# ----------------------------- PDF/Text extraction -----------------------------

def extract_pdf_pages(file_obj) -> List[Tuple[int, str]]:
    if file_obj is None:
        return []
    data = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    pages = []
    try:
        with fitz.open(stream=data, filetype="pdf") as doc:
            for i, page in enumerate(doc, 1):
                txt = page.get_text("text") or ""
                pages.append((i, txt))
    except Exception as e:
        st.warning(f"PDF extraction failed: {e}")
    return pages


def join_pages(pages: List[Tuple[int, str]]) -> str:
    return "\n".join([f"\n--- PAGE {p} ---\n{t}" for p, t in pages])


def lines_from_text(text: str) -> List[str]:
    return [clean_text(x) for x in text.splitlines() if clean_text(x)]


def window_after(lines: List[str], pattern: str, n: int = 8) -> str:
    rx = re.compile(pattern, re.I)
    for i, line in enumerate(lines):
        if rx.search(line):
            return " ".join(lines[i:i+n])
    return ""


def find_value_after_label(lines: List[str], label_regex: str, value_regex: str = r"(.+)", max_lines: int = 3) -> str:
    rx = re.compile(label_regex, re.I)
    for i, line in enumerate(lines):
        if rx.search(line):
            # Same line after label
            m = rx.search(line)
            rest = clean_text(line[m.end():])
            if rest:
                vm = re.search(value_regex, rest, re.I)
                if vm:
                    return clean_text(vm.group(1) if vm.groups() else vm.group(0))
            for j in range(1, max_lines + 1):
                if i + j < len(lines):
                    cand = clean_text(lines[i + j])
                    if cand and not re.match(r"^\d+(\.\d+)*\.?$", cand):
                        vm = re.search(value_regex, cand, re.I)
                        if vm:
                            return clean_text(vm.group(1) if vm.groups() else vm.group(0))
    return ""


def next_value_after(lines: List[str], label_regex: str, max_lines: int = 6) -> str:
    """Return first plausible value after a label line. Designed for HVPQ/Q88 line-per-cell PDFs."""
    rx = re.compile(label_regex, re.I)
    skip = re.compile(r"^(\d+|\d+(\.\d+)+\.?|yes/no|date issued|date expires|last annual|last intermediate|date of|endorsement)$", re.I)
    label_like = re.compile(r"^(general information|certificates|classification|ownership and operation|environmental information|publications|special survey|dry dock|in water survey|condition assessment|date of|name$|flag$|type of|if other|what is|does |has |is )", re.I)
    for i, line in enumerate(lines):
        if rx.search(line):
            # If value exists after colon on same line, take it.
            if ':' in line:
                tail = clean_text(line.split(':', 1)[1])
                if tail and not rx.fullmatch(tail):
                    return tail
            for j in range(1, max_lines + 1):
                if i + j >= len(lines):
                    break
                cand = clean_text(lines[i + j])
                if not cand:
                    continue
                if re.fullmatch(r"\d+(\.\d+)+\.?", cand):
                    break
                if skip.match(cand):
                    continue
                # allow No/Yes/NA as values
                if cand.lower() in ['yes', 'no', 'na', 'n/a', 'not applicable', 'nil', 'none']:
                    return cand
                # skip obvious labels/question text unless it has a date/number/value after same line
                if label_like.match(cand) and not DATE_RE.search(cand):
                    continue
                return cand
    return ''

def q88_blocks(lines: List[str]) -> Dict[str, List[str]]:
    blocks: Dict[str, List[str]] = {}
    current = None
    for line in lines:
        txt = line.strip()
        m = re.match(r"^(\d+\.\d+[a-z]?)(?:\s+(.*))?$", txt, re.I)
        if m:
            current = m.group(1).lower()
            blocks[current] = []
            rest = clean_text(m.group(2) or "")
            if rest:
                blocks[current].append(rest)
            continue
        if current:
            blocks[current].append(line)
    return blocks

def block_answer(block: List[str], skip_label_lines: int = 1, max_answer_lines: int = 8) -> str:
    vals = []
    # Usually first line is question label, answer begins after it.
    for cand in block[skip_label_lines:]:
        c = clean_text(cand)
        if not c:
            continue
        if re.match(r"^(Tel:|Fax:|Telex:|Email:|Web:|IMO:|Company IMO#)", c, re.I) and vals:
            break
        vals.append(c)
        if len(vals) >= max_answer_lines:
            break
    return clean_text(" ".join(vals))

def add_q88_block_field(fields: List[FieldRecord], blocks: Dict[str, List[str]], qno: str, source: str, field_id: str, label: str = ''):
    b = blocks.get(qno.lower(), [])
    if not b:
        return
    ans = block_answer(b)
    add_field(fields, source, field_id, ans, label=label, raw=" | ".join(b))

def parse_cert_rows_from_sequence(lines: List[str], source: str, cert_order: str = 'hvpq') -> List[FieldRecord]:
    """Parse certificate rows by partitioning between certificate labels, avoiding bleed into next row."""
    fields: List[FieldRecord] = []
    positions: List[Tuple[int, str]] = []
    for i, line in enumerate(lines):
        key = cert_key_from_label(line)
        if not key and re.search(r"certificate|u\.s\.|international|civil|safety|document|vessel general|wreck|financial", line, re.I):
            key = cert_key_from_label(line + " " + (lines[i+1] if i+1 < len(lines) else ""))
        if not key:
            continue
        # avoid generic headers where possible
        if normalize_key(line) in ["certificates", "certificate dates", "certificate description"]:
            continue
        positions.append((i, key))
    for idx, (pos, key) in enumerate(positions):
        next_pos = positions[idx + 1][0] if idx + 1 < len(positions) else min(len(lines), pos + 8)
        block_lines = lines[pos:next_pos]
        block = " ".join(block_lines)
        dates = extract_dates(block)
        if not dates:
            continue
        # Special rows with issue date only
        if key in ["tonnage"]:
            add_field(fields, source, f"cert.{key}.issue", dates[0], raw=block)
            continue
        # VGP in HVPQ often has issue date and an old second date, but Class Status usually has issue only. Keep both for review.
        if cert_order == 'q88':
            if len(dates) >= 1: add_field(fields, source, f"cert.{key}.issue", dates[0], raw=block)
            if len(dates) >= 4: add_field(fields, source, f"cert.{key}.expiry", dates[3], raw=block)
            elif len(dates) >= 2: add_field(fields, source, f"cert.{key}.expiry", dates[-1], raw=block)
            if len(dates) >= 2: add_field(fields, source, f"cert.{key}.last_annual", dates[1], raw=block)
            if len(dates) >= 3: add_field(fields, source, f"cert.{key}.last_intermediate", dates[2], raw=block)
        elif cert_order == 'class':
            if len(dates) >= 1: add_field(fields, source, f"cert.{key}.issue", dates[0], raw=block)
            if len(dates) >= 2: add_field(fields, source, f"cert.{key}.expiry", dates[1], raw=block)
        else:
            if len(dates) >= 1: add_field(fields, source, f"cert.{key}.issue", dates[0], raw=block)
            if len(dates) >= 2: add_field(fields, source, f"cert.{key}.expiry", dates[1], raw=block)
            if len(dates) >= 3: add_field(fields, source, f"cert.{key}.last_annual", dates[2], raw=block)
            if len(dates) >= 4: add_field(fields, source, f"cert.{key}.last_intermediate", dates[3], raw=block)
    return fields

# ----------------------------- Document-specific extractors -----------------------------

def extract_hvpq(pages: List[Tuple[int, str]]) -> List[FieldRecord]:
    source = "HVPQ"
    text = join_pages(pages)
    lines = lines_from_text(text)
    fields: List[FieldRecord] = []

    # Section-aware direct regex fixes for HVPQ line-per-cell layout
    m_owner = re.search(r"1\.3\.1\s+Registered Owner.*?\bName\s+([^\n]+)", text, re.I | re.S)
    if m_owner: add_field(fields, source, "owner.registered_owner", m_owner.group(1), raw=m_owner.group(0)[:500])
    m_pni = re.search(r"1\.1\.13\s+P and I Club.*?If other, then specify\s+([^\n]+)", text, re.I | re.S)
    if m_pni: add_field(fields, source, "insurance.pni_club", m_pni.group(1), raw=m_pni.group(0)[:500])
    m_mmsi = re.search(r"MMSI\) number\?\s+(\d{6,12})", text, re.I)
    if m_mmsi: add_field(fields, source, "vessel.mmsi", m_mmsi.group(1), raw=m_mmsi.group(0))

    # Header basics
    m = re.search(r"IMO/LR Number\s+(\d{7})\s+([A-Z0-9 '\-]+)", text, re.I)
    if m:
        add_field(fields, source, "vessel.imo", m.group(1), raw=m.group(0))
        add_field(fields, source, "vessel.name", clean_text(m.group(2)), raw=m.group(0))
    add_field(fields, source, "hvpq.date_completed", find_value_after_label(lines, r"Date this HVPQ document completed"), label="HVPQ date completed")
    add_field(fields, source, "vessel.flag", next_value_after(lines, r"^Flag$"))
    add_field(fields, source, "vessel.port_registry", next_value_after(lines, r"^Port of Registry$"))
    add_field(fields, source, "vessel.call_sign", next_value_after(lines, r"^Call sign$"))
    add_field(fields, source, "vessel.mmsi", next_value_after(lines, r"MMSI"))
    add_field(fields, source, "vessel.type", next_value_after(lines, r"type of ship as described"))
    other_type = next_value_after(lines, r"^If other, then specify$")
    if other_type and normalize_value(other_type) not in ["na", "not applicable"]:
        add_field(fields, source, "vessel.type_detail", other_type, label="Vessel type detail")
    add_field(fields, source, "owner.registered_owner", next_value_after(lines, r"^Name$"), raw="Registered owner name from 1.3.1")
    add_field(fields, source, "owner.technical_operator", next_value_after(lines, r"Technical operator"))
    add_field(fields, source, "owner.commercial_operator", next_value_after(lines, r"Commercial operator"))
    add_field(fields, source, "insurance.pni_club", find_value_after_label(lines, r"If other, then specify"), raw="P&I if-other extraction")
    # More specific P&I window
    pni_win = window_after(lines, r"P and I Club", 8)
    m = re.search(r"If other, then specify\s+(.+?)(?:\s+3\s+Amount|\s+Amount|$)", pni_win, re.I)
    if m:
        add_field(fields, source, "insurance.pni_club", m.group(1), raw=pni_win)

    # Environment
    cii_win = window_after(lines, r"Carbon Intensity Indicator|CII", 12)
    m = re.search(r"provide CII rating\s+([A-Z])\b", cii_win, re.I)
    if m: add_field(fields, source, "environment.cii_rating", m.group(1), raw=cii_win)
    m = re.search(r"CII rating verified by Class, 3rd Party or Owner\??\s+([A-Za-z0-9 /]+)", cii_win, re.I)
    if m: add_field(fields, source, "environment.cii_verified_by", m.group(1), raw=cii_win)
    else: add_field(fields, source, "environment.cii_verified_by", next_value_after(lines, r"CII rating verified by Class"), raw=cii_win)
    eexi_win = window_after(lines, r"EEXI|Energy Efficiency Existing", 12)
    m = re.search(r"provide EEXI rating\s+([0-9.]+)", eexi_win, re.I)
    if m: add_field(fields, source, "environment.eexi_rating", m.group(1), raw=eexi_win)
    m = re.search(r"EEXI rating verified by Class, 3rd Party or Owner\??\s+([A-Za-z0-9 /]+)", eexi_win, re.I)
    if m: add_field(fields, source, "environment.eexi_verified_by", m.group(1), raw=eexi_win)
    else: add_field(fields, source, "environment.eexi_verified_by", next_value_after(lines, r"EEXI rating verified by Class"), raw=eexi_win)

    # Class / survey
    add_field(fields, source, "classification.class_society", next_value_after(lines, r"^Classification Society$"))
    notation_win = window_after(lines, r"Class notation", 8)
    m = re.search(r"List class notations\s+(.+?)(?:Provide details|1\.5\.3|Change of Classification|$)", notation_win, re.I)
    if m: add_field(fields, source, "classification.class_notation", m.group(1), raw=notation_win)
    for fid, patt in [
        ("surveys.last_drydock", r"Date of last dry dock"),
        ("surveys.next_drydock_due", r"Date next dry dock due"),
        ("surveys.last_iws", r"Date of last IWS"),
        ("surveys.next_iws_due", r"Date next IWS due"),
        ("surveys.last_special", r"Date of last special survey"),
        ("surveys.next_special_due", r"Date next special survey due"),
        ("surveys.last_annual", r"Date of last annual survey"),
        ("surveys.last_intermediate", r"Date of Last Intermediate survey"),
    ]:
        add_field(fields, source, fid, next_value_after(lines, patt), raw=patt)
    add_field(fields, source, "classification.conditions_of_class", next_value_after(lines, r"Does Vessel have any open Conditions of Class"))
    add_field(fields, source, "classification.memo_of_class", next_value_after(lines, r"Does Vessel have any Memoranda of Class"))
    add_field(fields, source, "classification.flag_dispensation", next_value_after(lines, r"Does vessel have any flag state dispensations"))

    # Incidents / PSC
    add_field(fields, source, "incidents.pollution_grounding_collision_allision", next_value_after(lines, r"pollution, grounding, collision or allision"))
    add_field(fields, source, "incidents.other_incidents", next_value_after(lines, r"other incidents during the past 12 months"))
    add_field(fields, source, "psc.last_date", next_value_after(lines, r"Date of last Port State Control Inspection"))
    add_field(fields, source, "psc.last_port", next_value_after(lines, r"Port of last Port State Control Inspection"))
    add_field(fields, source, "psc.detained_36m", next_value_after(lines, r"detained during the last 36 months"))

    # Certificates: restrict extraction to the HVPQ Certificate dates table only.
    cert_start = next((i for i,l in enumerate(lines) if re.search(r"Certificate dates", l, re.I)), -1)
    cert_end = next((i for i,l in enumerate(lines[cert_start+1:], cert_start+1) if re.search(r"Publications", l, re.I)), len(lines)) if cert_start >= 0 else -1
    cert_lines = lines[cert_start:cert_end] if cert_start >= 0 else []
    fields += parse_cert_rows_from_sequence(cert_lines, source, cert_order="hvpq")
    return dedupe_fields(fields)


def extract_q88(pages: List[Tuple[int, str]]) -> List[FieldRecord]:
    source = "Q88"
    text = join_pages(pages)
    lines = lines_from_text(text)
    fields: List[FieldRecord] = []

    blocks = q88_blocks(lines)
    # Robust block-based extraction for Q88 table layout
    def ans(q): return block_answer(blocks.get(q, []))
    v = ans("1.1"); add_field(fields, source, "q88.date_updated", v, label="Q88 date updated", raw=" | ".join(blocks.get("1.1", [])))
    v = ans("1.2")
    m = re.search(r"(.+?)\s*\((\d{7})\)", v)
    if m:
        add_field(fields, source, "vessel.name", m.group(1), raw=v); add_field(fields, source, "vessel.imo", m.group(2), raw=v)
    v = ans("1.5")
    if "/" in v:
        a,b = v.split("/",1); add_field(fields, source, "vessel.flag", a, raw=v); add_field(fields, source, "vessel.port_registry", b, raw=v)
    v = ans("1.6")
    if "/" in v:
        a,b = v.split("/",1); add_field(fields, source, "vessel.call_sign", a, raw=v); add_field(fields, source, "vessel.mmsi", b, raw=v)
    for q, fid in [("1.8", "vessel.type"), ("1.8a", "vessel.type_detail"), ("1.10", "owner.registered_owner"), ("1.11", "owner.technical_operator"), ("1.12", "owner.commercial_operator"), ("1.14", "insurance.pni_club"), ("1.18", "classification.class_society"), ("1.19", "classification.class_notation")]:
        v = ans(q); add_field(fields, source, fid, v, raw=" | ".join(blocks.get(q, [])))
    for q, fid in [("1.20", "classification.conditions_of_class"), ("1.20a", "classification.memo_of_class")]:
        blocktxt = " | ".join(blocks.get(q, []))
        yns = re.findall(r"\b(Yes|No)\b", blocktxt, re.I)
        v = yns[-1] if yns else ans(q)
        add_field(fields, source, fid, v, raw=blocktxt)
    # survey dates
    for q in ["1.23", "1.24", "1.25", "1.25a"]:
        v = ans(q); dates = extract_dates(v)
        if q == "1.23" and dates: add_field(fields, source, "surveys.last_drydock", dates[0], raw=v)
        if q == "1.24" and dates:
            add_field(fields, source, "surveys.next_drydock_due", dates[0], raw=v)
            if len(dates)>1: add_field(fields, source, "surveys.next_annual_due", dates[1], raw=v)
        if q == "1.25" and dates:
            add_field(fields, source, "surveys.last_special", dates[0], raw=v)
            if len(dates)>1: add_field(fields, source, "surveys.next_special_due", dates[1], raw=v)
        if q == "1.25a" and dates:
            add_field(fields, source, "surveys.last_iws", dates[0], raw=v)
            if len(dates)>1: add_field(fields, source, "surveys.next_iws_due", dates[1], raw=v)
    # certificates in Q88 section 2
    for q, block in blocks.items():
        if re.match(r"2\.\d+", q):
            blocktxt = " ".join(block)
            key = cert_key_from_label(blocktxt)
            if key:
                dates = extract_dates(blocktxt)
                if len(dates)>=1: add_field(fields, source, f"cert.{key}.issue", dates[0], raw=blocktxt)
                if len(dates)>=4: add_field(fields, source, f"cert.{key}.expiry", dates[3], raw=blocktxt)
                elif len(dates)>=2: add_field(fields, source, f"cert.{key}.expiry", dates[-1], raw=blocktxt)
                if len(dates)>=2: add_field(fields, source, f"cert.{key}.last_annual", dates[1], raw=blocktxt)
                if len(dates)>=3: add_field(fields, source, f"cert.{key}.last_intermediate", dates[2], raw=blocktxt)
    # Q88 line-based extraction fallback
    for line in lines:
        if re.search(r"1\.1\s+Date updated", line, re.I):
            add_field(fields, source, "q88.date_updated", line.split(":")[-1], label="Q88 date updated", raw=line)
        m = re.search(r"1\.2\s+Vessel.*?:\s*(.+?)\s*\((\d{7})\)", line, re.I)
        if m:
            add_field(fields, source, "vessel.name", m.group(1), raw=line)
            add_field(fields, source, "vessel.imo", m.group(2), raw=line)
        m = re.search(r"1\.5\s+Flag/Port of Registry:\s*(.+?)/(.*)$", line, re.I)
        if m:
            add_field(fields, source, "vessel.flag", m.group(1), raw=line)
            add_field(fields, source, "vessel.port_registry", m.group(2), raw=line)
        m = re.search(r"1\.6\s+Call sign/MMSI:\s*([^/]+)/(.+)$", line, re.I)
        if m:
            add_field(fields, source, "vessel.call_sign", m.group(1), raw=line)
            add_field(fields, source, "vessel.mmsi", m.group(2), raw=line)
        m = re.search(r"1\.8\s+Type of vessel.*?:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "vessel.type", m.group(1), raw=line)
        m = re.search(r"1\.8a.*?specify:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "vessel.type_detail", m.group(1), label="Vessel type detail", raw=line)
        m = re.search(r"1\.10\s+Registered owner.*?:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "owner.registered_owner", m.group(1), raw=line)
        m = re.search(r"1\.11\s+Technical operator.*?:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "owner.technical_operator", m.group(1), raw=line)
        m = re.search(r"1\.12\s+Commercial operator.*?:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "owner.commercial_operator", m.group(1), raw=line)
        m = re.search(r"1\.14\s+P\s*&\s*I Club.*?:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "insurance.pni_club", m.group(1), raw=line)
        m = re.search(r"1\.18\s+Classification society:\s*(.+)$", line, re.I)
        if m: add_field(fields, source, "classification.class_society", m.group(1), raw=line)
        # conditions/memo handled by block parser above
        for fid, qno in [("surveys.last_drydock", "1.23"), ("surveys.next_drydock_due", "1.24"), ("surveys.last_special", "1.25"), ("surveys.last_iws", "1.25a")]:
            if line.startswith(qno):
                dates = extract_dates(line)
                if qno == "1.23" and dates:
                    add_field(fields, source, "surveys.last_drydock", dates[0], raw=line)
                elif qno == "1.24" and dates:
                    add_field(fields, source, "surveys.next_drydock_due", dates[0], raw=line)
                    if len(dates) > 1: add_field(fields, source, "surveys.next_annual_due", dates[1], raw=line)
                elif qno == "1.25" and dates:
                    add_field(fields, source, "surveys.last_special", dates[0], raw=line)
                    if len(dates) > 1: add_field(fields, source, "surveys.next_special_due", dates[1], raw=line)
                elif qno == "1.25a" and dates:
                    add_field(fields, source, "surveys.last_iws", dates[0], raw=line)
                    if len(dates) > 1: add_field(fields, source, "surveys.next_iws_due", dates[1], raw=line)
        # Certificate rows Q88 order: issued, last annual, last intermediate, expires
        if re.match(r"^2\.\d+\s+", line):
            key = cert_key_from_label(line)
            if key:
                dates = extract_dates(line)
                if len(dates) >= 1: add_field(fields, source, f"cert.{key}.issue", dates[0], raw=line)
                if len(dates) >= 4: add_field(fields, source, f"cert.{key}.expiry", dates[-1], raw=line)
                elif len(dates) >= 2: add_field(fields, source, f"cert.{key}.expiry", dates[-1], raw=line)
                if len(dates) >= 2: add_field(fields, source, f"cert.{key}.last_annual", dates[1], raw=line)
                if len(dates) >= 3: add_field(fields, source, f"cert.{key}.last_intermediate", dates[2], raw=line)
        m = re.search(r"CII rating.*?\b([A-E])\b", line, re.I)
        if m:
            add_field(fields, source, "environment.cii_rating", m.group(1), raw=line)
        if re.search(r"CII rating verified", line, re.I):
            tail = line.split("?")[-1].strip()
            if tail and tail.lower() != line.lower():
                add_field(fields, source, "environment.cii_verified_by", tail, raw=line)

    # Certificate special: P&I coverage/expiration contains expiry only
    for line in lines:
        m = re.search(r"1\.15\s+P\s*&\s*I Club pollution liability coverage/expiration date:\s*(.+)$", line, re.I)
        if m:
            dates = extract_dates(line)
            if dates:
                add_field(fields, source, "cert.pni_cover.expiry", dates[-1], raw=line)
    # Class notation window
    win = window_after(lines, r"1\.19\s+Class notation", 5)
    if win:
        add_field(fields, source, "classification.class_notation", re.sub(r"^1\.19\s+Class notation:\s*", "", win, flags=re.I), raw=win)

    return dedupe_fields(fields)


def extract_piq(pages: List[Tuple[int, str]]) -> List[FieldRecord]:
    source = "PIQ"
    text = join_pages(pages)
    lines = lines_from_text(text)
    fields: List[FieldRecord] = []
    m = re.search(r"Vessel Name\s+([A-Z0-9 '\-]+)\s+Date\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})", text, re.I)
    if m:
        add_field(fields, source, "vessel.name", m.group(1), raw=m.group(0))
        add_field(fields, source, "piq.date", m.group(2), label="PIQ date", raw=m.group(0))
    add_field(fields, source, "vessel.type", find_value_after_label(lines, r"^Vessel Type\b"))
    ann = find_value_after_label(lines, r"Annex II cargo")
    add_field(fields, source, "piq.annex_ii_carried_or_intended", ann, label="Annex II carried/intended")
    add_field(fields, source, "classification.last_class_visit", find_value_after_label(lines, r"Date of last visit"))
    add_field(fields, source, "classification.last_class_visit_purpose", find_value_after_label(lines, r"Purpose of visit"))

    # Superintendent visit blocks
    tech_win = window_after(lines, r"Technical Superintendent inspection completed", 25)
    mar_win = window_after(lines, r"Marine Superintendent inspection completed", 25)
    add_visit_fields(fields, source, "technical", tech_win)
    add_visit_fields(fields, source, "marine", mar_win)

    # Tank inspections
    tank_patterns = [
        ("tank.cargo_slop.freq_months", "tank.cargo_slop.oldest", r"cargo tanks", r"cargo and slop"),
        ("tank.ballast.freq_months", "tank.ballast.oldest", r"ballast tanks", r"ballast tanks"),
        ("tank.void.freq_months", "tank.void.oldest", r"void spaces", r"void spaces"),
    ]
    for fid_freq, fid_old, p1, p2 in tank_patterns:
        win = window_after(lines, p1, 8)
        m = re.search(r"Required frequency.*?(\d+)\s+months", win, re.I)
        if m: add_field(fields, source, fid_freq, m.group(1), label=fid_freq, raw=win)
        dates = extract_dates(win)
        if dates: add_field(fields, source, fid_old, dates[-1], label=fid_old, raw=win)

    # CAP / MOC
    add_field(fields, source, "cap.enrolled", find_value_after_label(lines, r"Enrolled in a condition assessment programme"))
    add_field(fields, source, "cap.overall_rating", find_value_after_label(lines, r"Overall rating"))
    add_field(fields, source, "cap.completed", find_value_after_label(lines, r"CAP survey completed"))
    add_field(fields, source, "cap.class_society", find_value_after_label(lines, r"Which classification society issued"))
    add_field(fields, source, "moc.structural_change", find_value_after_label(lines, r"Have any structural changes been made"))
    add_field(fields, source, "moc.retrofit", find_value_after_label(lines, r"Has any new equipment been retrofitted"))
    retrofit_win = window_after(lines, r"Has any new equipment been retrofitted", 10)
    add_field(fields, source, "moc.retrofit_details", retrofit_win, label="Retrofit details", raw=retrofit_win)
    add_field(fields, source, "moc.equipment_replaced", find_value_after_label(lines, r"Equipment replaced"))
    add_field(fields, source, "moc.equipment_replaced_details", window_after(lines, r"Equipment replaced", 8), label="Equipment replaced details")
    add_field(fields, source, "moc.equipment_decommissioned", find_value_after_label(lines, r"Equipment decommissioned"))

    # PSC block
    psc_win = window_after(lines, r"last three Port State Control", 20)
    add_field(fields, source, "psc.block", psc_win, label="PSC block", raw=psc_win)
    m = re.search(r"Last\s+" + DATE_RE.pattern + r"\s+([^,]+)", psc_win, re.I)
    if m:
        dates = extract_dates(m.group(0))
        if dates: add_field(fields, source, "psc.last_date", dates[0], raw=psc_win)
        add_field(fields, source, "psc.last_port", m.group(1), raw=psc_win)

    # Incidents: PIQ incident section often later, extract yes/no if visible
    for kw, fid in [
        ("pollution", "incidents.pollution"), ("grounding", "incidents.grounding"),
        ("collision", "incidents.collision"), ("allision", "incidents.allision"),
        ("mooring", "incidents.mooring"), ("injury", "incidents.injury"),
        ("machinery", "incidents.machinery"),
    ]:
        win = window_after(lines, kw, 4)
        m = re.search(r"\b(Yes|No)\b", win, re.I)
        if m: add_field(fields, source, fid, m.group(1), label=f"PIQ incident {kw}", raw=win)
    # collapsed incident yes/no for no incidents declared logic
    no_count = sum(1 for f in fields if f.field_id.startswith("incidents.") and normalize_bool(f.value) == "no")
    yes_count = sum(1 for f in fields if f.field_id.startswith("incidents.") and normalize_bool(f.value) == "yes")
    if yes_count == 0 and no_count > 0:
        add_field(fields, source, "incidents.other_incidents", "No", raw="Derived from PIQ incident section")

    return dedupe_fields(fields)


def add_visit_fields(fields: List[FieldRecord], source: str, kind: str, win: str):
    if not win:
        return
    add_field(fields, source, f"superintendent.{kind}.raw", win, label=f"{kind.title()} superintendent raw", raw=win)
    # Rows look like Last 01 October 2025 28 November 2025 59.00 Physical
    rows = re.findall(r"(Second Last|Third Last|Last)\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+([0-9.]+)?\s*([A-Za-z]*)", win, re.I)
    for idx, row in enumerate(rows, 1):
        label, frm, to, days, typ = row
        add_field(fields, source, f"superintendent.{kind}.{idx}.from", frm, label=f"{kind.title()} superintendent {label} from", raw=win)
        add_field(fields, source, f"superintendent.{kind}.{idx}.to", to, label=f"{kind.title()} superintendent {label} to", raw=win)
        if days: add_field(fields, source, f"superintendent.{kind}.{idx}.days", days, label=f"{kind.title()} superintendent {label} days", raw=win)
        if typ: add_field(fields, source, f"superintendent.{kind}.{idx}.type", typ, label=f"{kind.title()} superintendent {label} type", raw=win)


def extract_class_status(pages: List[Tuple[int, str]]) -> List[FieldRecord]:
    source = "CLASS"
    text = join_pages(pages)
    lines = lines_from_text(text)
    fields: List[FieldRecord] = []
    if re.search(r"KOREAN REGISTER|\bKR\b", text, re.I):
        class_soc = "Korean Register"
    elif re.search(r"NIPPON KAIJI KYOKAI|ClassNK|NK-SHIPS", text, re.I):
        class_soc = "Nippon Kaiji Kyokai"
    else:
        class_soc = find_value_after_label(lines, r"Class Society|Classification Society")
    add_field(fields, source, "classification.class_society", class_soc)
    add_field(fields, source, "vessel.name", next_value_after(lines, r"^Ship Name$|^Name of Ship$"))
    m_imo = re.search(r"IMO No\s*:?\s*(\d{7})|IMO Number\s*:?\s*(\d{7})", text, re.I)
    add_field(fields, source, "vessel.imo", (m_imo.group(1) or m_imo.group(2)) if m_imo else next_value_after(lines, r"^IMO No\.?$|^IMO Number$"))
    add_field(fields, source, "vessel.flag", next_value_after(lines, r"^Flag$"))
    add_field(fields, source, "vessel.port_registry", next_value_after(lines, r"^Port of Registry$"))
    add_field(fields, source, "classification.class_notation", window_after(lines, r"Class Notation|Classification Character", 6))
    add_field(fields, source, "owner.registered_owner", next_value_after(lines, r"^Owner$|^Registered Owner$"))
    add_field(fields, source, "owner.technical_operator", next_value_after(lines, r"^Technical Manager$|^Management Company$"))

    # KR/ClassNK certificate tables: restrict to certificate description section
    cstart = next((i for i,l in enumerate(lines) if re.search(r"Certificate description", l, re.I)), -1)
    cend = next((i for i,l in enumerate(lines[cstart+1:], cstart+1) if re.search(r"Survey Description", l, re.I)), len(lines)) if cstart >= 0 else -1
    class_cert_lines = lines[cstart:cend] if cstart >= 0 else []
    fields += parse_cert_rows_from_sequence(class_cert_lines, source, cert_order="class")

    # KR certificate table sometimes has certificate desc rows followed by issue/expiry dates. Scan all lines.
    for line in lines:
        key = cert_key_from_label(line)
        if key:
            dates = extract_dates(line)
            # KR order usually issue/expiry or expiry only depending page. Use text labels when possible.
            if dates:
                if re.search(r"Expiry", line, re.I) and len(dates) == 1:
                    add_field(fields, source, f"cert.{key}.expiry", dates[0], raw=line)
                elif len(dates) >= 2:
                    add_field(fields, source, f"cert.{key}.issue", dates[0], raw=line)
                    add_field(fields, source, f"cert.{key}.expiry", dates[-1], raw=line)
                elif len(dates) == 1:
                    # for ClassNK Current Statutory Certificates rows list only expiry date
                    add_field(fields, source, f"cert.{key}.expiry", dates[0], raw=line)
    # ClassNK Current Statutory Certificates: row has cert + expiry date
    for i, line in enumerate(lines):
        key = cert_key_from_label(line)
        if key and not sources_value(fields, f"cert.{key}.expiry").get(source):
            win = " ".join(lines[i:i+3])
            dates = extract_dates(win)
            if dates:
                add_field(fields, source, f"cert.{key}.expiry", dates[-1], raw=win)

    # Survey status rows KR/NK
    for line in lines:
        norm = normalize_key(line)
        dates = extract_dates(line)
        if not dates:
            continue
        if "special survey" in norm:
            if re.search(r"next|due", line, re.I): add_field(fields, source, "surveys.next_special_due", dates[0], raw=line)
            add_field(fields, source, "surveys.last_special", dates[-1], raw=line)
        if "annual survey" in norm:
            if re.search(r"due|range", line, re.I): add_field(fields, source, "surveys.next_annual_due", dates[0], raw=line)
            add_field(fields, source, "surveys.last_annual", dates[-1], raw=line)
        if "intermediate survey" in norm:
            add_field(fields, source, "surveys.last_intermediate", dates[-1], raw=line)
        if "docking survey" in norm or "dock" in norm:
            if re.search(r"next|due", line, re.I): add_field(fields, source, "surveys.next_drydock_due", dates[0], raw=line)
            add_field(fields, source, "surveys.last_drydock", dates[-1], raw=line)
        if "iws" in norm or "in water" in norm:
            if re.search(r"next|due", line, re.I): add_field(fields, source, "surveys.next_iws_due", dates[0], raw=line)
            add_field(fields, source, "surveys.last_iws", dates[-1], raw=line)

    # Conditions / memo / recommendations: be conservative, don't confuse empty sections with mismatch.
    cc_text = " ".join([l for l in lines if re.search(r"condition(s)? of class|recommendation|memoranda|memorandum|memo", l, re.I)])
    if cc_text:
        if re.search(r"no\s+(open\s+)?condition|none|nil", cc_text, re.I):
            add_field(fields, source, "classification.conditions_of_class", "No", raw=cc_text)
        elif re.search(r"condition", cc_text, re.I):
            add_field(fields, source, "classification.conditions_of_class", "Review listed items", raw=cc_text)
        if re.search(r"no\s+(memoranda|memorandum|memo)|none|nil", cc_text, re.I):
            add_field(fields, source, "classification.memo_of_class", "No", raw=cc_text)

    return dedupe_fields(fields)


def extract_xml(file_obj) -> List[FieldRecord]:
    source = "XML"
    fields: List[FieldRecord] = []
    if file_obj is None:
        return fields
    data = file_obj.getvalue() if hasattr(file_obj, "getvalue") else file_obj.read()
    try:
        root = ET.fromstring(data)
        ns = {"x": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
        vessel = root.find(".//x:Vessel", ns) if ns else root.find(".//Vessel")
        if vessel is not None:
            add_field(fields, source, "vessel.name", vessel.attrib.get("name", ""), raw=ET.tostring(vessel, encoding="unicode"))
            add_field(fields, source, "vessel.imo", vessel.attrib.get("id", ""), raw=ET.tostring(vessel, encoding="unicode"))
        doc = root.find(".//x:Document", ns) if ns else root.find(".//Document")
        if doc is not None:
            add_field(fields, source, "xml.exported", doc.attrib.get("exported", ""), label="XML exported date")
        templ = root.find(".//x:Template", ns) if ns else root.find(".//Template")
        if templ is not None:
            add_field(fields, source, "xml.template", templ.attrib.get("variant", ""), label="XML template")
    except Exception as e:
        add_field(fields, source, "xml.error", str(e), label="XML parse error")
    return fields


def dedupe_fields(fields: List[FieldRecord]) -> List[FieldRecord]:
    seen = set()
    out = []
    for f in fields:
        k = (f.source, f.field_id, normalize_value(f.value), f.date_value)
        if k in seen:
            continue
        seen.add(k)
        out.append(f)
    return out

# ----------------------------- Optional local LLM extraction assist -----------------------------

def ollama_generate(base_url: str, model: str, prompt: str, timeout: int = 120) -> str:
    url = base_url.rstrip("/") + "/api/generate"
    payload = {"model": model, "prompt": prompt, "stream": False, "options": {"temperature": 0.0}}
    r = requests.post(url, json=payload, timeout=timeout)
    r.raise_for_status()
    return r.json().get("response", "")


def extract_json_from_llm_response(resp: str) -> Dict[str, Any]:
    resp = resp.strip()
    m = re.search(r"\{.*\}", resp, re.S)
    if not m:
        return {}
    try:
        return json.loads(m.group(0))
    except Exception:
        return {}


def llm_assist_extract(doc_type: str, pages: List[Tuple[int, str]], base_url: str, model: str) -> List[FieldRecord]:
    source = doc_type.upper() + "_LLM"
    text = join_pages(pages)
    # Use only relevant first 18k chars to avoid local context overflow; class/Q88 certs in early pages generally.
    snippet = text[:18000]
    prompt = f"""
You are a marine vetting document extraction engine. Extract only values visible in the text. Do not guess.
Document type: {doc_type}
Return ONLY valid JSON. If value is not visible, use null.

Required JSON keys:
{{
  "vessel": {{"name": null, "imo": null, "flag": null, "port_registry": null, "type": null, "call_sign": null, "mmsi": null}},
  "owner": {{"registered_owner": null, "technical_operator": null, "commercial_operator": null}},
  "insurance": {{"pni_club": null}},
  "classification": {{"class_society": null, "class_notation": null, "conditions_of_class": null, "memo_of_class": null, "flag_dispensation": null}},
  "environment": {{"cii_rating": null, "cii_verified_by": null, "eexi_rating": null, "eexi_verified_by": null}},
  "surveys": {{"last_drydock": null, "next_drydock_due": null, "last_iws": null, "next_iws_due": null, "last_special": null, "next_special_due": null, "last_annual": null, "next_annual_due": null, "last_intermediate": null}},
  "certificates": {{
    "cofr": {{"issue": null, "expiry": null}}, "iopp": {{"issue": null, "expiry": null}}, "vgp": {{"issue": null, "expiry": null}},
    "doc": {{"issue": null, "expiry": null, "last_annual": null}}, "smc": {{"issue": null, "expiry": null, "last_annual": null}},
    "issc": {{"issue": null, "expiry": null, "last_annual": null}}, "class_certificate": {{"issue": null, "expiry": null, "last_annual": null, "last_intermediate": null}},
    "safety_equipment": {{"issue": null, "expiry": null, "last_annual": null, "last_intermediate": null}},
    "safety_radio": {{"issue": null, "expiry": null, "last_annual": null, "last_intermediate": null}},
    "safety_construction": {{"issue": null, "expiry": null, "last_annual": null, "last_intermediate": null}},
    "loadline": {{"issue": null, "expiry": null, "last_annual": null, "last_intermediate": null}},
    "cof_chemical": {{"issue": null, "expiry": null, "last_annual": null, "last_intermediate": null}}
  }},
  "incidents": {{"pollution_grounding_collision_allision": null, "other_incidents": null}}
}}

TEXT:
{snippet}
"""
    try:
        resp = ollama_generate(base_url, model, prompt)
        data = extract_json_from_llm_response(resp)
    except Exception as e:
        return [FieldRecord(source=source, field_id="llm.error", label="LLM extraction error", value=str(e), confidence="llm-error")]
    fields: List[FieldRecord] = []
    flatten_llm_json(data, source, fields)
    return fields


def flatten_llm_json(data: Dict[str, Any], source: str, fields: List[FieldRecord], prefix: str = ""):
    for k, v in (data or {}).items():
        fid = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            if prefix == "certificates":
                cert_key = k
                for sub, val in v.items():
                    if val not in [None, "", []]:
                        add_field(fields, source, f"cert.{cert_key}.{sub}", val, confidence="local-llm")
            else:
                flatten_llm_json(v, source, fields, fid)
        elif v not in [None, "", []]:
            add_field(fields, source, fid, v, confidence="local-llm")

# ----------------------------- Rule engine -----------------------------

def add_finding(findings: List[Finding], **kwargs):
    findings.append(Finding(**kwargs))


def compare_field(findings: List[Finding], fields: List[FieldRecord], field_id: str, sources: List[str], area: str, check_name: str, risk: str = "MEDIUM", hard: bool = False):
    vals = {s: first_field(fields, s, field_id) for s in sources}
    vals = {s: v for s, v in vals.items() if clean_text(v)}
    if len(vals) < 2:
        return
    pairs = list(vals.items())
    base_s, base_v = pairs[0]
    mismatch = False
    for s, v in pairs[1:]:
        if not semantically_equivalent(base_v, v, field_id):
            mismatch = True
    if mismatch:
        status = "MISMATCH" if hard else "MANUAL CHECK"
        add_finding(findings, area=area, check=check_name, status=status, risk=risk,
                    hvpq_value=vals.get("HVPQ", ""), piq_value=vals.get("PIQ", ""), class_value=vals.get("CLASS", ""), q88_value=vals.get("Q88", ""), xml_value=vals.get("XML", ""),
                    reason="Mapped values differ after normalization. Treat as manual check unless the source document confirms exact official value.",
                    action="Verify source documents and update the stale/incorrect declaration.")


def run_rules(fields: List[FieldRecord], ref_date: date, settings: Dict[str, Any], obs_df: pd.DataFrame) -> List[Finding]:
    findings: List[Finding] = []

    # 1. Identity only big mismatches
    compare_field(findings, fields, "vessel.imo", ["HVPQ", "PIQ", "Q88", "CLASS", "XML"], "Identity", "IMO consistency across documents", "CRITICAL", hard=True)
    compare_field(findings, fields, "vessel.name", ["HVPQ", "PIQ", "Q88", "CLASS", "XML"], "Identity", "Vessel name consistency across documents", "HIGH", hard=False)

    # 2. Vessel type as manual consistency, not hard mismatch
    vals_type = sources_value(fields, "vessel.type")
    if len(vals_type) >= 2:
        # Only flag as manual if textual difference visible; not high mismatch.
        normset = set(normalize_value(v) for v in vals_type.values() if v)
        if len(normset) > 1:
            add_finding(findings, area="General", check="Vessel type wording consistency", status="MANUAL CHECK", risk="MEDIUM",
                        hvpq_value=vals_type.get("HVPQ", ""), piq_value=vals_type.get("PIQ", ""), class_value=vals_type.get("CLASS", ""), q88_value=vals_type.get("Q88", ""),
                        reason="Vessel type wording differs across documents. This may be acceptable for oil/product/chemical wording but should be confirmed against IOPP/COF.",
                        action="Confirm official type wording and Annex II/chemical carriage applicability.")

    # 3. Owner/P&I spelling/wording
    compare_field(findings, fields, "owner.registered_owner", ["HVPQ", "Q88", "CLASS"], "Ownership", "Registered owner consistency", "MEDIUM", hard=False)
    vals_pni = sources_value(fields, "insurance.pni_club")
    if len(vals_pni) >= 2:
        hv, qv = vals_pni.get("HVPQ", ""), vals_pni.get("Q88", "")
        if hv and qv and not semantically_equivalent(hv, qv, "insurance.pni_club"):
            risk = "LOW" if similarity(hv, qv) > 0.72 else "MEDIUM"
            add_finding(findings, area="Insurance", check="P&I club spelling / consistency", status="MANUAL CHECK", risk=risk,
                        hvpq_value=hv, q88_value=qv, reason="P&I club name differs or appears misspelled.", action="Correct typo/style in HVPQ/Q88 if required.")

    # 4. Class society broad, official class status only useful here
    compare_field(findings, fields, "classification.class_society", ["HVPQ", "Q88", "CLASS"], "Class", "Class society consistency", "HIGH", hard=False)
    for fid, label in [("classification.conditions_of_class", "Conditions of Class"), ("classification.memo_of_class", "Memoranda of Class"), ("classification.flag_dispensation", "Flag/Class dispensation")]:
        vals = sources_value(fields, fid)
        if len(vals) >= 2:
            bools = {s: normalize_bool(v) for s, v in vals.items()}
            if "yes" in bools.values() and "no" in bools.values():
                add_finding(findings, area="Class", check=f"{label} declaration mismatch", status="MISMATCH", risk="CRITICAL",
                            hvpq_value=vals.get("HVPQ", ""), class_value=vals.get("CLASS", ""), q88_value=vals.get("Q88", ""),
                            reason=f"{label} appears differently declared across documents.", action="Verify latest Class Status and update HVPQ/Q88.")

    # 5. Superintendent gaps
    run_superintendent_rules(findings, fields, ref_date)

    # 6. Certificates: compare HVPQ/Q88/CLASS mapped fields.
    certs = sorted(set([f.field_id.split(".")[1] for f in fields if f.field_id.startswith("cert.") and len(f.field_id.split(".")) >= 3]))
    for cert in certs:
        if cert in ["tonnage"]:
            continue
        for part in ["issue", "expiry"]:
            if cert == "vgp" and part == "expiry":
                continue
            fid = f"cert.{cert}.{part}"
            vals = sources_value(fields, fid)
            if len(vals) < 2:
                continue
            # Prefer Class/Q88 as current cross-check sources for certs. If HVPQ differs from any by date, flag.
            hv = vals.get("HVPQ", "")
            qv = vals.get("Q88", "")
            cv = vals.get("CLASS", "")
            hvd, qvd, cvd = parse_date_any(hv), parse_date_any(qv), parse_date_any(cv)
            source_newer = qvd or cvd
            # Expiry critical if HVPQ expired but other source valid later
            if part == "expiry" and hvd:
                latest = max([d for d in [qvd, cvd] if d] or [hvd])
                if hvd < ref_date <= latest and latest > hvd:
                    add_finding(findings, area="Certificates", check=f"{cert.upper()} expiry outdated/expired in HVPQ", status="MISMATCH", risk="CRITICAL",
                                hvpq_value=hv, q88_value=qv, class_value=cv,
                                reason="HVPQ certificate expiry appears expired or older while another source shows a later valid expiry.",
                                action="Verify latest certificate and update HVPQ immediately.")
                    continue
            # Issue/expiry mismatches larger than tolerance.
            dlist = [("HVPQ", hvd), ("Q88", qvd), ("CLASS", cvd)]
            dlist = [(s, d) for s, d in dlist if d]
            if len(dlist) >= 2:
                dates_only = [d for s, d in dlist]
                if max(dates_only) != min(dates_only):
                    day_gap = (max(dates_only) - min(dates_only)).days
                    if day_gap > (0 if part == "expiry" else 7):
                        risk = "HIGH"
                        if cert in ["cofr", "iopp", "smc", "issc", "class_certificate"] and part == "expiry":
                            risk = "CRITICAL" if min(dates_only) < ref_date else "HIGH"
                        add_finding(findings, area="Certificates", check=f"{cert.upper()} {part} date mismatch", status="MANUAL CHECK", risk=risk,
                                    hvpq_value=hv, q88_value=qv, class_value=cv,
                                    reason=f"Certificate {part} dates differ across mapped sources by {day_gap} days.",
                                    action="Check the latest certificate/Class Status and correct stale declarations.")

    # 7. Certificate sanity: annual before issue etc.
    for src in ["HVPQ", "Q88"]:
        for cert in certs:
            issue = parse_date_any(first_field(fields, src, f"cert.{cert}.issue"))
            ann = parse_date_any(first_field(fields, src, f"cert.{cert}.last_annual"))
            exp = parse_date_any(first_field(fields, src, f"cert.{cert}.expiry"))
            if issue and ann and ann < issue - relativedelta(days=1):
                add_finding(findings, area="Certificates", check=f"{cert.upper()} annual endorsement before issue date", status="MANUAL CHECK", risk="HIGH",
                            hvpq_value=first_field(fields, src, f"cert.{cert}.issue") + " / " + first_field(fields, src, f"cert.{cert}.last_annual") if src == "HVPQ" else "",
                            q88_value=first_field(fields, src, f"cert.{cert}.issue") + " / " + first_field(fields, src, f"cert.{cert}.last_annual") if src == "Q88" else "",
                            reason="Last annual/endorsement date appears earlier than certificate issue date; may be table extraction issue or stale entry.", action="Verify certificate row manually.")
            other_expiries = []
            for osrc in ["HVPQ", "Q88", "CLASS"]:
                if osrc != src:
                    od = parse_date_any(first_field(fields, osrc, f"cert.{cert}.expiry"))
                    if od: other_expiries.append(od)
            if exp and exp < ref_date and cert not in ["vgp", "tonnage"] and not any(od > exp for od in other_expiries):
                add_finding(findings, area="Certificates", check=f"{cert.upper()} appears expired in {src}", status="MISMATCH", risk="CRITICAL",
                            hvpq_value=first_field(fields, src, f"cert.{cert}.expiry") if src == "HVPQ" else "",
                            q88_value=first_field(fields, src, f"cert.{cert}.expiry") if src == "Q88" else "",
                            reason=f"Certificate expiry in {src} is before reference date {ref_date.isoformat()}.", action="Update expired/stale certificate data.")

    # 8. Surveys: IWS next due blank if last IWS exists
    for src in ["HVPQ", "Q88", "CLASS"]:
        last = first_field(fields, src, "surveys.last_iws")
        nxt = first_field(fields, src, "surveys.next_iws_due")
        if last and not nxt:
            add_finding(findings, area="Class / Survey", check=f"IWS next due missing in {src}", status="MANUAL CHECK", risk="HIGH",
                        hvpq_value=last if src == "HVPQ" else "", q88_value=last if src == "Q88" else "", class_value=last if src == "CLASS" else "",
                        reason="Last IWS is declared but next IWS due is blank/not extracted.", action="Confirm whether IWS remains applicable and update HVPQ/Q88 if required.")
    for fid, label in [("surveys.last_drydock", "Last dry dock"), ("surveys.next_drydock_due", "Next dry dock due"), ("surveys.last_special", "Last special survey"), ("surveys.next_special_due", "Next special survey due")]:
        compare_dates_mapped(findings, fields, fid, label)

    # 9. CII/EEXI blank/mismatch checks
    vals = sources_value(fields, "environment.cii_rating")
    if len(vals) >= 2 and len(set(normalize_value(v) for v in vals.values())) > 1:
        add_finding(findings, area="Environmental", check="CII rating mismatch", status="MISMATCH", risk="HIGH",
                    hvpq_value=vals.get("HVPQ", ""), q88_value=vals.get("Q88", ""), reason="CII rating differs across documents.", action="Verify latest CII statement/SEEMP and update.")
    hv_cii_basis = first_field(fields, "HVPQ", "environment.cii_verified_by")
    q_cii_basis = first_field(fields, "Q88", "environment.cii_verified_by")
    q_cii_rating = first_field(fields, "Q88", "environment.cii_rating")
    if q_cii_rating and hv_cii_basis and not q_cii_basis:
        add_finding(findings, area="Environmental", check="Q88 CII verification basis blank", status="MANUAL CHECK", risk="MEDIUM",
                    hvpq_value=hv_cii_basis, q88_value=q_cii_basis,
                    reason="Q88 appears to contain CII rating but verification basis is blank/not extracted while HVPQ has a basis.", action="Verify Q88 CII verification basis.")

    # 10. Incidents
    hv_inc = normalize_bool(first_field(fields, "HVPQ", "incidents.other_incidents"))
    hv_pgca = normalize_bool(first_field(fields, "HVPQ", "incidents.pollution_grounding_collision_allision"))
    piq_inc = normalize_bool(first_field(fields, "PIQ", "incidents.other_incidents"))
    q_inc = normalize_bool(first_field(fields, "Q88", "incidents.other_incidents"))
    # No incidents across docs -> manual positive confirmation
    if all(x in ["", "no"] for x in [hv_inc, hv_pgca, piq_inc, q_inc]) and any(x == "no" for x in [hv_inc, hv_pgca, piq_inc, q_inc]):
        add_finding(findings, area="Incidents", check="No incidents declared", status="MANUAL CHECK", risk="MEDIUM",
                    hvpq_value=first_field(fields, "HVPQ", "incidents.other_incidents") or first_field(fields, "HVPQ", "incidents.pollution_grounding_collision_allision"),
                    piq_value=first_field(fields, "PIQ", "incidents.other_incidents"), q88_value=first_field(fields, "Q88", "incidents.other_incidents"),
                    reason="No incidents appear declared. This is not a mismatch but needs positive confirmation from vessel/office.",
                    action="Confirm no reportable machinery, navigation, mooring, pollution, security, injury or operational incidents in the last 12 months.")
    if "yes" in [hv_inc, hv_pgca, piq_inc, q_inc] and "no" in [hv_inc, hv_pgca, piq_inc, q_inc]:
        add_finding(findings, area="Incidents", check="Incident declaration mismatch", status="MISMATCH", risk="CRITICAL",
                    hvpq_value=str(hv_inc or hv_pgca), piq_value=str(piq_inc), q88_value=str(q_inc),
                    reason="One document indicates incident(s) while another indicates no incidents.", action="Verify incident register and update PIQ/HVPQ/Q88.")

    # 11. PIQ tank inspections due checks
    for name, old_fid, freq_fid in [("Cargo/slop", "tank.cargo_slop.oldest", "tank.cargo_slop.freq_months"), ("Ballast", "tank.ballast.oldest", "tank.ballast.freq_months"), ("Void space", "tank.void.oldest", "tank.void.freq_months")]:
        old = parse_date_any(first_field(fields, "PIQ", old_fid))
        freq = first_field(fields, "PIQ", freq_fid)
        try:
            months = int(float(freq)) if freq else 12
        except Exception:
            months = 12
        if old:
            due = old + relativedelta(months=months)
            days = (due - ref_date).days
            risk = "HIGH" if days < 0 else ("MEDIUM" if days <= 60 else "LOW")
            status = "MISMATCH" if days < 0 else "MANUAL CHECK"
            if risk != "LOW" or settings.get("show_low", False):
                add_finding(findings, area="Tank Inspection", check=f"{name} tank inspection sequence due calculation", status=status, risk=risk,
                            piq_value=f"Oldest date: {old.isoformat()}, frequency: {months} months, due: {due.isoformat()}",
                            reason="PIQ tank inspection sequence due date calculated from oldest inspection date and required frequency.",
                            action="Vessel to confirm full tank inspection sequence is complete/current and records are available.")

    # 12. MOC/retrofit checklist
    if normalize_bool(first_field(fields, "PIQ", "moc.retrofit")) == "yes" or normalize_bool(first_field(fields, "PIQ", "moc.equipment_replaced")) == "yes":
        add_finding(findings, area="MOC / Retrofit", check="Retrofit/replacement declared in PIQ", status="MANUAL CHECK", risk="MEDIUM",
                    piq_value=(first_field(fields, "PIQ", "moc.retrofit_details") + " " + first_field(fields, "PIQ", "moc.equipment_replaced_details"))[:500],
                    reason="PIQ declares retrofit/replacement. HVPQ/Q88/Class/certificates should reflect any affected statutory/class entries.",
                    action="Confirm MOC close-out, updated certificates/forms, class records and HVPQ/Q88 declarations.")

    # 13. Required blank/missing checks: HVPQ/Q88 core fields blank
    required_hvpq = [
        "vessel.name", "vessel.imo", "vessel.flag", "vessel.type", "classification.class_society",
        "surveys.last_drydock", "surveys.next_drydock_due", "surveys.last_special", "surveys.next_special_due",
        "psc.last_date", "psc.detained_36m", "environment.cii_rating", "environment.cii_verified_by",
        "insurance.pni_club", "classification.conditions_of_class", "classification.memo_of_class",
    ]
    for fid in required_hvpq:
        if not first_field(fields, "HVPQ", fid):
            add_finding(findings, area="Blank / Missing", check=f"HVPQ missing {FIELD_LABELS.get(fid, fid)}", status="MANUAL CHECK", risk="MEDIUM",
                        hvpq_value="blank", reason="Required/commonly observed HVPQ field is blank or not extracted.", action="Verify and complete if applicable.")

    # 14. Observation-library driven checklist, not mismatch
    for area, action in observation_checklist_from_excel(obs_df):
        add_finding(findings, area="Observation Library", check=area, status="MANUAL CHECK", risk="MEDIUM", reason="Recurring historical observation pattern found in uploaded observation library.", action=action)

    return dedupe_findings(findings)


def compare_dates_mapped(findings: List[Finding], fields: List[FieldRecord], fid: str, label: str):
    vals = sources_value(fields, fid)
    if len(vals) < 2:
        return
    parsed = {s: parse_date_any(v) for s, v in vals.items() if parse_date_any(v)}
    if len(parsed) < 2:
        return
    if max(parsed.values()) != min(parsed.values()):
        day_gap = (max(parsed.values()) - min(parsed.values())).days
        if day_gap > 7:
            add_finding(findings, area="Class / Survey", check=f"{label} date mismatch", status="MANUAL CHECK", risk="HIGH",
                        hvpq_value=vals.get("HVPQ", ""), q88_value=vals.get("Q88", ""), class_value=vals.get("CLASS", ""),
                        reason=f"Mapped survey dates differ across sources by {day_gap} days.", action="Verify latest class status/certificate and update stale declaration.")


def run_superintendent_rules(findings: List[Finding], fields: List[FieldRecord], ref_date: date):
    for kind, max_days, label in [("technical", MONTHS_7_DAYS, "Technical Superintendent"), ("marine", MONTHS_12_DAYS, "Marine Superintendent")]:
        visits = []
        for i in range(1, 5):
            frm = parse_date_any(first_field(fields, "PIQ", f"superintendent.{kind}.{i}.from"))
            to = parse_date_any(first_field(fields, "PIQ", f"superintendent.{kind}.{i}.to"))
            typ = first_field(fields, "PIQ", f"superintendent.{kind}.{i}.type")
            if frm and to:
                visits.append((frm, to, typ))
        visits = sorted(visits, key=lambda x: x[0])
        for idx in range(len(visits) - 1):
            prev_to = visits[idx][1]
            next_from = visits[idx + 1][0]
            gap = (next_from - prev_to).days
            if gap > max_days:
                add_finding(findings, area="Management Oversight", check=f"{label} inspection gap exceeds limit", status="MISMATCH", risk="CRITICAL",
                            piq_value=f"Previous to: {prev_to.isoformat()}, next from: {next_from.isoformat()}, gap: {gap} days",
                            reason=f"{label} inspection gap exceeds locked rule ({'7' if kind=='technical' else '12'} months, strict no tolerance).",
                            action="Office/vessel to provide justification and update inspection plan; flag for pre-vetting review.")
        if visits:
            last_to = visits[-1][1]
            gap = (ref_date - last_to).days
            if gap > max_days:
                add_finding(findings, area="Management Oversight", check=f"{label} last inspection overdue as of reference date", status="MISMATCH", risk="CRITICAL",
                            piq_value=f"Last to: {last_to.isoformat()}, reference: {ref_date.isoformat()}, gap: {gap} days",
                            reason=f"{label} last inspection exceeds locked interval as of reference date.", action="Arrange inspection/verify updated PIQ.")
        else:
            add_finding(findings, area="Management Oversight", check=f"{label} inspection dates not extracted", status="MANUAL CHECK", risk="HIGH",
                        piq_value="Not extracted", reason="PIQ says management oversight should be verified but visit rows were not extracted.", action="Check PIQ superintendent visit table manually.")


def dedupe_findings(findings: List[Finding]) -> List[Finding]:
    seen = set()
    out = []
    priority = {"CRITICAL": 0, "HIGH": 1, "MEDIUM": 2, "LOW": 3}
    for f in findings:
        k = (f.area, f.check, f.hvpq_value, f.piq_value, f.class_value, f.q88_value)
        if k in seen:
            continue
        seen.add(k)
        out.append(f)
    out.sort(key=lambda x: (priority.get(x.risk.upper(), 9), x.area, x.check))
    return out

# ----------------------------- Observation library -----------------------------

def parse_obs_excel(file_obj) -> pd.DataFrame:
    if file_obj is None:
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(file_obj)
        frames = []
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df["_sheet"] = sheet
            frames.append(df)
        if not frames:
            return pd.DataFrame()
        df = pd.concat(frames, ignore_index=True)
        df["joined"] = df.fillna("").astype(str).apply(lambda r: " ".join([str(v) for v in r.tolist()]), axis=1)
        return df
    except Exception as e:
        return pd.DataFrame({"joined": [f"Observation Excel parse error: {e}"]})


def observation_checklist_from_excel(obs_df: pd.DataFrame) -> List[Tuple[str, str]]:
    if obs_df is None or obs_df.empty or "joined" not in obs_df.columns:
        return []
    text = "\n".join(obs_df["joined"].astype(str).tolist()).lower()
    checks = []
    patterns = [
        ("Mooring ropes / brake test / split drums", ["mooring", "brake", "rope", "split drum", "rendering"], "Verify mooring winch brake test dates/BHC/rendering load, rope/tail particulars, end-for-end/retirement records, and split drum declarations."),
        ("Cargo/ballast/void tank inspection records", ["tank", "coating", "ballast", "void"], "Verify tank coating/inspection dates, frequency, condition and that all cargo/slop/ballast/void spaces are covered."),
        ("Certificate dates / endorsements", ["certificate", "expiry", "annual", "endorsement"], "Verify certificate issue/expiry/endorsement dates against latest certificates/Class Status."),
        ("PSC history", ["psc", "port state", "deficien"], "Verify last three PSC dates/ports/MOU/deficiencies/detention and OCIMF PSC database entry."),
        ("CII/EEXI/EEDI", ["cii", "eexi", "eedi", "aer"], "Verify latest energy efficiency values, rating and verification basis."),
        ("Publications", ["publication", "edition", "isgott", "meg", "colreg", "mfag"], "Verify listed publication editions against onboard publication list/SMS requirement."),
        ("Firefighting / foam", ["foam", "fire", "fixed firefighting", "sample locker"], "Verify foam type/test date, fixed firefighting systems and sample locker arrangements."),
        ("Pollution prevention / overboard blanks", ["overboard", "scupper", "sea chest", "pressure test", "bunker piping", "cargo piping"], "Verify overboard discharge blanks/testing arrangement, sea chest wording, scupper plugs, and cargo/bunker pressure test records."),
        ("Cargo / IGS / venting", ["igs", "inert gas", "vent", "p/v", "vapour", "tank gauging", "cow"], "Verify IGS, venting/PV valve, vapour return, tank gauging and COW declarations."),
        ("Diagrams / arrangements", ["diagram", "manifold", "fairlead", "chock", "bitt", "bow mooring"], "Verify mooring/manifold/fairlead/chock/bitt/bow arrangement diagrams are attached/current."),
        ("Lifting gear / cranes", ["crane", "lifting", "swl", "hoist"], "Verify lifting gear/crane annual and five-year tests, SWL and certificates."),
    ]
    for title, keys, action in patterns:
        if any(k in text for k in keys):
            checks.append((title, action))
    return checks

# ----------------------------- Export helpers -----------------------------

def df_from_fields(fields: List[FieldRecord]) -> pd.DataFrame:
    return pd.DataFrame([asdict(f) for f in fields]) if fields else pd.DataFrame(columns=["source", "field_id", "label", "value", "date_value", "confidence", "raw"])


def df_from_findings(findings: List[Finding]) -> pd.DataFrame:
    return pd.DataFrame([asdict(f) for f in findings]) if findings else pd.DataFrame(columns=["area", "check", "status", "risk", "hvpq_value", "piq_value", "class_value", "q88_value", "xml_value", "reason", "action"])


def make_excel(findings: List[Finding], fields: List[FieldRecord], checklist: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_from_findings(findings).to_excel(writer, index=False, sheet_name="Findings")
        df_from_fields(fields).to_excel(writer, index=False, sheet_name="Extracted Fields")
        checklist.to_excel(writer, index=False, sheet_name="Ship Checklist")
    return bio.getvalue()


def make_ship_checklist(findings: List[Finding]) -> pd.DataFrame:
    rows = []
    for f in findings:
        if f.risk.upper() in ["CRITICAL", "HIGH", "MEDIUM"]:
            rows.append({
                "Area": f.area,
                "Check": f.check,
                "Risk": f.risk,
                "What vessel/office should verify": f.action or f.reason,
                "HVPQ": f.hvpq_value,
                "PIQ": f.piq_value,
                "Class Status": f.class_value,
                "Q88": f.q88_value,
            })
    return pd.DataFrame(rows)

# ----------------------------- Streamlit app -----------------------------

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    with st.sidebar:
        st.header("Upload documents")
        hvpq_file = st.file_uploader("HVPQ PDF", type=["pdf"], key="hvpq")
        hvpq_xml = st.file_uploader("HVPQ XML (optional)", type=["xml"], key="xml")
        piq_file = st.file_uploader("PIQ PDF", type=["pdf"], key="piq")
        q88_file = st.file_uploader("Q88 PDF", type=["pdf"], key="q88")
        class_file = st.file_uploader("Class Status PDF", type=["pdf"], key="class")
        obs_file = st.file_uploader("HVPQ observation library Excel", type=["xlsx", "xls"], key="obs")
        inc_obs_file = st.file_uploader("Incident observation library Excel", type=["xlsx", "xls"], key="incobs")
        st.divider()
        ref_date_input = st.date_input("Reference / review date", value=date.today())
        show_low = st.checkbox("Show low-risk findings", value=False)
        use_llm = st.checkbox("Use local LLM extraction assist (Ollama)", value=False)
        ollama_url = st.text_input("Ollama URL", value="http://localhost:11434", disabled=not use_llm)
        ollama_model = st.text_input("Ollama model", value="qwen2.5:14b", disabled=not use_llm)
        run_btn = st.button("Run extraction and checks", type="primary")

    if not run_btn:
        st.info("Upload documents and click **Run extraction and checks**. For best accuracy, review/edit extracted fields before relying on findings.")
        return

    settings = {"show_low": show_low}
    all_fields: List[FieldRecord] = []
    page_cache = {}

    with st.spinner("Extracting text and structured fields..."):
        if hvpq_file:
            pages = extract_pdf_pages(hvpq_file); page_cache["HVPQ"] = pages
            all_fields += extract_hvpq(pages)
            if use_llm: all_fields += llm_assist_extract("HVPQ", pages, ollama_url, ollama_model)
        if piq_file:
            pages = extract_pdf_pages(piq_file); page_cache["PIQ"] = pages
            all_fields += extract_piq(pages)
            if use_llm: all_fields += llm_assist_extract("PIQ", pages, ollama_url, ollama_model)
        if q88_file:
            pages = extract_pdf_pages(q88_file); page_cache["Q88"] = pages
            all_fields += extract_q88(pages)
            if use_llm: all_fields += llm_assist_extract("Q88", pages, ollama_url, ollama_model)
        if class_file:
            pages = extract_pdf_pages(class_file); page_cache["CLASS"] = pages
            all_fields += extract_class_status(pages)
            if use_llm: all_fields += llm_assist_extract("CLASS", pages, ollama_url, ollama_model)
        if hvpq_xml:
            all_fields += extract_xml(hvpq_xml)
        obs_df = pd.concat([parse_obs_excel(obs_file), parse_obs_excel(inc_obs_file)], ignore_index=True) if (obs_file or inc_obs_file) else pd.DataFrame()
        all_fields = dedupe_fields(all_fields)

    st.subheader("Step 1 — Review extracted fields")
    st.caption("Critical point: findings are only as good as extracted values. Edit any wrong extracted values below, then rerun checks using the edited table.")
    fields_df = df_from_fields(all_fields)
    edited_df = st.data_editor(fields_df, num_rows="dynamic", use_container_width=True, height=360)
    edited_fields = []
    for _, r in edited_df.fillna("").iterrows():
        edited_fields.append(FieldRecord(source=str(r.get("source", "")), field_id=str(r.get("field_id", "")), label=str(r.get("label", "")), value=str(r.get("value", "")), date_value=str(r.get("date_value", "")), confidence=str(r.get("confidence", "")), raw=str(r.get("raw", ""))))

    findings = run_rules(edited_fields, ref_date_input, settings, obs_df)
    if not show_low:
        findings = [f for f in findings if f.risk.upper() != "LOW"]
    findings_df = df_from_findings(findings)
    checklist_df = make_ship_checklist(findings)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Actionable rows", len(findings_df))
    c2.metric("Critical", int((findings_df.get("risk", pd.Series(dtype=str)).str.upper() == "CRITICAL").sum()) if not findings_df.empty else 0)
    c3.metric("High", int((findings_df.get("risk", pd.Series(dtype=str)).str.upper() == "HIGH").sum()) if not findings_df.empty else 0)
    c4.metric("Manual/Medium", int((findings_df.get("risk", pd.Series(dtype=str)).str.upper() == "MEDIUM").sum()) if not findings_df.empty else 0)

    tabs = st.tabs(["Findings", "Ship checklist", "Extracted fields", "Observation library", "Raw text debug", "JSON export"])
    with tabs[0]:
        st.subheader("Actionable mismatch / manual-check register")
        st.dataframe(findings_df, use_container_width=True, height=500)
    with tabs[1]:
        st.subheader("Simple ship/office verification checklist")
        st.dataframe(checklist_df, use_container_width=True, height=500)
    with tabs[2]:
        st.subheader("Extracted structured fields")
        st.dataframe(edited_df, use_container_width=True, height=500)
    with tabs[3]:
        st.subheader("Observation library patterns")
        if obs_df.empty:
            st.info("No observation Excel uploaded.")
        else:
            st.dataframe(obs_df.head(500), use_container_width=True, height=500)
    with tabs[4]:
        st.subheader("Raw text debug")
        src = st.selectbox("Source", list(page_cache.keys()) or ["None"])
        if src != "None":
            page_no = st.number_input("Page", min_value=1, max_value=max([p for p, _ in page_cache[src]]), value=1)
            txt = dict(page_cache[src]).get(page_no, "")
            st.text_area("Extracted page text", txt, height=500)
    with tabs[5]:
        st.subheader("JSON export")
        st.json({"findings": [asdict(f) for f in findings], "fields": [asdict(f) for f in edited_fields]})

    xlsx = make_excel(findings, edited_fields, checklist_df)
    st.download_button("Download Excel register", xlsx, file_name="hvpq_piq_q88_class_verification_register.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
