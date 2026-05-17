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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from dateutil import parser as dateparser
from dateutil.relativedelta import relativedelta

APP_TITLE = "HVPQ / PIQ / Q88 / Class Status Checker v13"
APP_SUBTITLE = "Simple review dashboard with HVPQ correction register, vessel action list, Q88 value-add checks, and manual-confirmation coverage."

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
    # User rule: if tank inspection dates are within the required cycle (normally 12 months), do NOT flag.
    # Valid/current tank inspections are reported positively in the Office Summary instead.
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
            if days < 0:
                add_finding(findings, area="Tank Inspection", check=f"{name} tank inspection overdue", status="MISMATCH", risk="HIGH",
                            piq_value=f"Oldest date: {old.isoformat()}, frequency: {months} months, due: {due.isoformat()}",
                            reason="PIQ tank inspection sequence due date is overdue based on oldest inspection date and required frequency.",
                            action="Vessel to confirm completed inspection sequence and update PIQ/supporting records immediately.")
        elif freq:
            add_finding(findings, area="Tank Inspection", check=f"{name} tank oldest inspection date missing", status="MANUAL CHECK", risk="MEDIUM",
                        piq_value=f"Frequency: {months} months; oldest inspection date not extracted",
                        reason="PIQ shows tank inspection frequency but oldest inspection date was blank/not extracted.",
                        action="Vessel to confirm the oldest inspection date in the current sequence and provide records.")

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
                        hvpq_value="blank", reason="Required/commonly observed HVPQ field is blank or not extracted.", action="Verify and complete HVPQ if applicable. HVPQ is the main document to correct.")

    # 13b. Mapped blanks in PIQ and Q88. These are not automatic defects; they are included so no blank/uncertain field is silently missed.
    has_piq = any(f.source == "PIQ" for f in fields)
    has_q88 = any(f.source == "Q88" for f in fields)
    if has_piq:
        required_piq = [
            "vessel.name", "vessel.type", "psc.last_date", "psc.detained_36m",
            "superintendent.technical.1.from", "superintendent.technical.1.to",
            "superintendent.marine.1.from", "superintendent.marine.1.to",
            "tank.cargo_slop.oldest", "tank.ballast.oldest", "tank.void.oldest",
        ]
        for fid in required_piq:
            if not first_field(fields, "PIQ", fid):
                add_finding(findings, area="Blank / Missing", check=f"PIQ missing/not extracted {FIELD_LABELS.get(fid, fid)}", status="MANUAL CHECK", risk="MEDIUM",
                            piq_value="blank/not extracted", reason="Mapped PIQ field is blank or could not be reliably extracted.",
                            action="Verify the PIQ entry manually and correct PIQ if blank/stale/wrong.")
    if has_q88:
        required_q88 = [
            "vessel.name", "vessel.imo", "vessel.type", "classification.class_society",
            "classification.conditions_of_class", "classification.memo_of_class",
            "surveys.last_drydock", "surveys.next_drydock_due",
            "surveys.last_special", "surveys.next_special_due",
            "environment.cii_rating", "environment.cii_verified_by",
            "insurance.pni_club",
        ]
        for fid in required_q88:
            if not first_field(fields, "Q88", fid):
                add_finding(findings, area="Blank / Missing", check=f"Q88 missing/not extracted {FIELD_LABELS.get(fid, fid)}", status="MANUAL CHECK", risk="MEDIUM",
                            q88_value="blank/not extracted", reason="Mapped Q88 value-add field is blank or could not be reliably extracted.",
                            action="Verify Q88 entry manually. If Q88 disagrees with HVPQ, use source evidence/Class Status before changing HVPQ.")

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


def make_excel(findings: List[Finding], fields: List[FieldRecord], office_report: pd.DataFrame, vessel_actions: pd.DataFrame, obs_pinpoints: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        office_report.to_excel(writer, index=False, sheet_name="Office Summary")
        vessel_actions.to_excel(writer, index=False, sheet_name="Vessel Action Checklist")
        df_from_findings(findings).to_excel(writer, index=False, sheet_name="Detailed Findings")
        obs_pinpoints.to_excel(writer, index=False, sheet_name="Obs Library Pinpoints")
        df_from_fields(fields).to_excel(writer, index=False, sheet_name="Extracted Fields")
        # Light formatting: freeze header row and set column widths.
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = min(max((len(str(cell.value)) if cell.value is not None else 0) for cell in col) + 2, 70)
                ws.column_dimensions[col[0].column_letter].width = max(12, max_len)
    return bio.getvalue()



# ----------------------------- Office summary and vessel action helpers -----------------------------

QID_ACTION_MAP = {
    "1.1.8": ("Vessel type", "Confirm type of ship wording against Form A/B of IOPP/IOPPC and Q88. Ensure oil/product/chemical wording is consistent and Annex II carriage is correctly declared."),
    "1.1.9": ("Vessel type other-specify", "If HVPQ says 'Other', confirm the free-text vessel type is completed exactly as per certificate/Q88."),
    "1.2.1": ("EEDI", "Verify EEDI applicability, value/reason for exemption and verification basis against latest energy efficiency documentation."),
    "1.2.2": ("EEXI", "Verify EEXI rating/value and verification basis against latest EEXI technical file/certificate."),
    "1.2.3": ("CII/AER", "Verify latest CII rating, AER value and verification basis. If Q88/HVPQ blank or inconsistent, update before vetting."),
    "1.3.1": ("Registered owner", "Confirm registered owner full style, address, country, contact and IMO number against certificate/Q88/Class Status."),
    "1.3.2": ("Technical operator", "Confirm technical operator full style, company IMO number, DPA and emergency contacts against Q88/SMS."),
    "1.3.3": ("Commercial operator", "Confirm commercial/disponent operator details against current chartering data/Q88."),
    "1.4.3": ("Building contract date", "Confirm building contract date against builder/class records; correct HVPQ if date differs."),
    "1.5.1": ("Class society", "Confirm class society name and IACS status against current Class Status."),
    "1.5.2": ("Class notation", "Confirm full class notation is current and includes relevant notations such as IWS, IGS, COW, BWT, EGCS, LG etc. where applicable."),
    "1.5.4": ("Dry dock", "Check last/second-last/next dry dock dates and location against Class Status."),
    "1.5.5": ("IWS", "Check last IWS and next IWS due date. If next due is blank, confirm whether IWS remains applicable or was reset by drydock/renewal."),
    "1.5.6": ("Special survey", "Check last/next special survey dates and ESS/IACS details against Class Status."),
    "1.5.10": ("Thickness measurement", "Check date of last thickness measurement against class/CAP records."),
    "1.5.11": ("Annual survey", "Check last annual survey and next annual survey window against Class Status."),
    "1.5.12": ("Intermediate survey", "Check last intermediate survey date against Class Status."),
    "1.5.14": ("Conditions of Class", "Confirm open Conditions of Class / recommendations exactly as per Class Status; if none, ensure all docs say No/Nil."),
    "1.5.16": ("Memoranda of Class", "Confirm memoranda/notes/recommendations exactly as per Class Status; if none, ensure all docs say No/Nil."),
    "1.5.18": ("Flag/Class dispensation", "Check any flag/class dispensations or exemptions and ensure HVPQ/Q88 match current documents."),
    "1.9.2": ("Unscheduled repairs", "Confirm whether unscheduled repairs since last special survey/drydock are declared where required."),
    "1.9.3": ("Pollution/grounding/collision/allision", "Confirm no pollution, grounding, collision or allision incident in last 12 months, or declare details if any."),
    "1.9.5": ("Other incidents", "Confirm no other reportable incidents in last 12 months, including machinery, injury, mooring, security or operational incidents."),
    "1.9.6": ("Incident details", "If incidents exist, verify date/type/location/details are complete and consistent with PIQ/Q88."),
    "1.9.8": ("PSC", "Verify last PSC date, port, MOU, deficiencies, detention and OCIMF PSC database entry."),
    "2.1.5": ("Certificate dates", "Verify certificate issue/expiry/annual/intermediate/endorsement dates against latest certificates and Class Status."),
    "2.2.1": ("Publications", "Verify onboard publication editions against latest publication list/SMS and update HVPQ if stale."),
    "5.3.1": ("Fixed foam", "Verify foam type, test analysis certificate date and cargo-area fixed foam details."),
    "5.3.2": ("Fixed firefighting systems", "Verify fixed firefighting systems for paint locker, pump room, ER, lockers and other spaces."),
    "6.1.8": ("Sea chest", "Verify sea chest/cargo piping segregation wording and whether cargo sea chest is fitted/not provided."),
    "6.1.10": ("Overboard blanks/testing", "Verify overboard discharge blanks or testing arrangement and supporting records."),
    "6.1.13": ("Cargo piping pressure test", "Verify cargo piping pressure test policy, pressure and latest test record."),
    "6.1.14": ("Bunker piping pressure test", "Verify bunker piping pressure test policy, pressure and latest test record."),
    "7.1.1": ("Cargo tank coating", "Verify all cargo/slop/residual tanks are listed with coating type, extent, condition and latest inspection date."),
    "7.1.3": ("Ballast tank coating", "Verify all ballast tanks are listed with coating type, extent, condition and latest inspection date."),
    "7.1.5": ("Structural inspection programme", "Verify formal inspection programme covers void, hold, cargo and ballast spaces."),
    "10.1.4": ("Mooring winch / brake test", "Verify mooring winch particulars, split drum declaration, BHC/rendering load and latest brake test dates against records."),
    "10.1.7": ("Mooring ropes", "Verify mooring rope/tail particulars, certificates, end-for-end dates, retirement/discard criteria and records."),
    "10.2.1": ("Fairlead/chock/bitt diagram", "Verify fairlead/chock/bitt diagram is attached/current and matches deck arrangement."),
    "10.7.1": ("Bow mooring arrangement", "Verify bow mooring arrangement diagram and equipment details."),
    "10.8.1": ("Manifold arrangement", "Verify manifold arrangement diagram, dimensions and reducers/spools match vessel."),
    "10.9.1": ("Lifting gear / cranes", "Verify crane/lifting gear SWL, annual test and five-year load test dates/certificates."),
}

AREA_QIDS = {
    "Mooring ropes / brake test / split drums": ["10.1.4", "10.1.7"],
    "Cargo/ballast/void tank inspection records": ["7.1.1", "7.1.3", "7.1.5"],
    "Certificate dates / endorsements": ["2.1.5"],
    "PSC history": ["1.9.8"],
    "CII/EEXI/EEDI": ["1.2.1", "1.2.2", "1.2.3"],
    "Publications": ["2.2.1"],
    "Firefighting / foam": ["5.3.1", "5.3.2"],
    "Pollution prevention / overboard blanks": ["6.1.8", "6.1.10", "6.1.13", "6.1.14"],
    "Diagrams / arrangements": ["10.2.1", "10.7.1", "10.8.1"],
    "Lifting gear / cranes": ["10.9.1"],
}

def obs_qids_from_text(text: str) -> List[str]:
    text = clean_text(text)
    # Keep only HVPQ-looking IDs; avoid dates and decimals by requiring at least two dots OR known one-dot fields.
    candidates = re.findall(r"\b(?:item\s*)?(\d{1,2}\.\d{1,4}(?:\.\d{1,4})?)\b", text, flags=re.I)
    out = []
    for q in candidates:
        # Filter obvious decimals/versions; accept known qids or those beginning with likely HVPQ chapters.
        if q in QID_ACTION_MAP or re.match(r"^(1|2|3|4|5|6|7|8|9|10|11|12)\.", q):
            if q not in out:
                out.append(q)
    return out


def observation_pinpoint_rows(obs_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if obs_df is None or obs_df.empty or "joined" not in obs_df.columns:
        return pd.DataFrame(columns=["Priority", "Question No.", "Topic", "What vessel should check", "Observation basis / example"])
    for _, r in obs_df.iterrows():
        joined = str(r.get("joined", ""))
        header_qids = obs_qids_from_text(" ".join([str(c) for c in obs_df.columns]))
        qids = obs_qids_from_text(joined)
        for q in header_qids + qids:
            if q in QID_ACTION_MAP:
                topic, action = QID_ACTION_MAP[q]
                rows.append({
                    "Priority": "Targeted",
                    "Question No.": q,
                    "Topic": topic,
                    "What vessel should check": action,
                    "Observation basis / example": clean_text(joined)[:450],
                })
    # Add generic area-to-question fallbacks when the observation library mentions the area but exact qid is not captured.
    text = "\n".join(obs_df["joined"].astype(str).tolist()).lower()
    for area, qids in AREA_QIDS.items():
        if area.lower().split(" /")[0] in text or any(k in text for k in normalize_key(area).split()[:2]):
            for q in qids:
                if q in QID_ACTION_MAP:
                    topic, action = QID_ACTION_MAP[q]
                    rows.append({"Priority": "Targeted", "Question No.": q, "Topic": topic, "What vessel should check": action, "Observation basis / example": f"Recurring observation library area: {area}"})
    if not rows:
        return pd.DataFrame(columns=["Priority", "Question No.", "Topic", "What vessel should check", "Observation basis / example"])
    df = pd.DataFrame(rows)
    return df.drop_duplicates(subset=["Question No.", "Topic", "What vessel should check"]).sort_values(["Question No.", "Topic"])


def build_positive_office_checks(fields: List[FieldRecord], findings: List[Finding], ref_date: date) -> pd.DataFrame:
    rows = []
    def has_bad(area_contains: str, check_contains: str = "") -> bool:
        for f in findings:
            if area_contains.lower() in f.area.lower() and check_contains.lower() in f.check.lower() and f.risk.upper() in ["CRITICAL", "HIGH"]:
                return True
        return False
    # Identity
    if not has_bad("Identity"):
        name = first_field(fields, "HVPQ", "vessel.name") or first_field(fields, "PIQ", "vessel.name") or first_field(fields, "Q88", "vessel.name") or first_field(fields, "CLASS", "vessel.name")
        imo = first_field(fields, "HVPQ", "vessel.imo") or first_field(fields, "XML", "vessel.imo") or first_field(fields, "Q88", "vessel.imo")
        if name or imo:
            rows.append({"Section": "Identity", "Status": "In order", "Office summary": f"Vessel identity broadly matches across uploaded documents: {name} / IMO {imo}."})
    # Tank inspections: positive if within cycle
    for title, old_fid, freq_fid in [("Cargo/slop tank inspections", "tank.cargo_slop.oldest", "tank.cargo_slop.freq_months"), ("Ballast tank inspections", "tank.ballast.oldest", "tank.ballast.freq_months"), ("Void space inspections", "tank.void.oldest", "tank.void.freq_months")]:
        old = parse_date_any(first_field(fields, "PIQ", old_fid))
        freq = first_field(fields, "PIQ", freq_fid)
        months = 12
        try:
            months = int(float(freq)) if freq else 12
        except Exception:
            pass
        if old:
            due = old + relativedelta(months=months)
            if due >= ref_date:
                rows.append({"Section": "Tank inspection", "Status": "In order", "Office summary": f"{title}: oldest date {old.isoformat()}, frequency {months} months, next due {due.isoformat()} — within required cycle."})
    # Superintendent
    if not has_bad("Management Oversight", "Technical"):
        v = first_field(fields, "PIQ", "superintendent.technical.1.to") or first_field(fields, "PIQ", "superintendent.technical.1.from")
        if v:
            rows.append({"Section": "Management oversight", "Status": "In order / no high-risk gap detected", "Office summary": f"Technical Superintendent visit data extracted; no high-risk gap detected by locked 7-month rule. Latest extracted date: {v}."})
    if not has_bad("Management Oversight", "Marine"):
        v = first_field(fields, "PIQ", "superintendent.marine.1.to") or first_field(fields, "PIQ", "superintendent.marine.1.from")
        if v:
            rows.append({"Section": "Management oversight", "Status": "In order / no high-risk gap detected", "Office summary": f"Marine Superintendent visit data extracted; no high-risk gap detected by locked 12-month rule. Latest extracted date: {v}."})
    # Certificates validity
    cert_exp = []
    for f in fields:
        if f.field_id.startswith("cert.") and f.field_id.endswith(".expiry"):
            d = parse_date_any(f.value)
            if d:
                cert_exp.append((f.field_id.split(".")[1], f.source, d))
    expired = [(c,s,d) for c,s,d in cert_exp if d < ref_date]
    high_cert_issues = [f.check for f in findings if f.area == "Certificates" and f.risk.upper() in ["CRITICAL", "HIGH"]]
    if cert_exp:
        if not high_cert_issues and not expired:
            rows.append({"Section": "Certificates", "Status": "In order", "Office summary": f"Extracted certificate expiry dates are valid as of {ref_date.isoformat()}; no high-risk certificate issue detected."})
        else:
            rows.append({"Section": "Certificates", "Status": "Exceptions noted", "Office summary": "Certificate review completed; exceptions are listed in the issue summary/vessel tab."})
    # Class conditions / memo
    coc_vals = sources_value(fields, "classification.conditions_of_class")
    memo_vals = sources_value(fields, "classification.memo_of_class")
    if coc_vals and all(normalize_bool(v) in ["no", "", "na"] for v in coc_vals.values()):
        rows.append({"Section": "Class", "Status": "In order", "Office summary": "Conditions of Class appear declared as nil/no in extracted sources."})
    if memo_vals and all(normalize_bool(v) in ["no", "", "na"] for v in memo_vals.values()):
        rows.append({"Section": "Class", "Status": "In order", "Office summary": "Memoranda of Class appear declared as nil/no in extracted sources."})
    return pd.DataFrame(rows, columns=["Section", "Status", "Office summary"])


def build_issue_summary(findings: List[Finding]) -> pd.DataFrame:
    rows = []
    for f in findings:
        if f.risk.upper() in ["CRITICAL", "HIGH"]:
            rows.append({
                "Priority": f.risk,
                "Area": f.area,
                "Issue": f.check,
                "Brief office note": f.reason,
                "Action": f.action,
            })
    return pd.DataFrame(rows, columns=["Priority", "Area", "Issue", "Brief office note", "Action"])


def make_vessel_action_checklist(findings: List[Finding], obs_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for f in findings:
        if f.risk.upper() in ["CRITICAL", "HIGH", "MEDIUM"]:
            qno = ""
            # Infer common qnos from area/check.
            check_norm = normalize_key(f.check + " " + f.area)
            if "cofr" in check_norm or "certificate" in check_norm or f.area == "Certificates": qno = "2.1.5"
            elif "technical superintendent" in check_norm or "marine superintendent" in check_norm: qno = "PIQ 2.2.1001 / 2.2.1002"
            elif "iws" in check_norm: qno = "1.5.5"
            elif "special survey" in check_norm: qno = "1.5.6"
            elif "dry dock" in check_norm: qno = "1.5.4"
            elif "cii" in check_norm: qno = "1.2.3"
            elif "p&i" in check_norm or "pni" in check_norm: qno = "1.1.13"
            elif "owner" in check_norm: qno = "1.3.1"
            elif "incident" in check_norm: qno = "1.9.3 / 1.9.5 / PIQ 5.7"
            elif "tank" in check_norm: qno = "2.3.3001-3003 / 7.1.1 / 7.1.3"
            elif "retrofit" in check_norm or "moc" in check_norm: qno = "PIQ 2.5.1002-1004"
            rows.append({
                "Priority": f.risk,
                "Question / Section": qno,
                "Area": f.area,
                "What to check": f.check,
                "Why flagged": f.reason,
                "Action requested from vessel/office": f.action,
                "HVPQ value": f.hvpq_value,
                "PIQ value": f.piq_value,
                "Class Status value": f.class_value,
                "Q88 value": f.q88_value,
            })
    obs_pin = observation_pinpoint_rows(obs_df)
    for _, r in obs_pin.iterrows():
        rows.append({
            "Priority": "Targeted check",
            "Question / Section": r.get("Question No.", ""),
            "Area": r.get("Topic", "Observation library"),
            "What to check": r.get("What vessel should check", ""),
            "Why flagged": "Recurring historical HVPQ/PIQ observation pattern. This is not a confirmed defect; it is a targeted pre-vetting check.",
            "Action requested from vessel/office": r.get("What vessel should check", ""),
            "HVPQ value": "", "PIQ value": "", "Class Status value": "", "Q88 value": "",
        })
    if not rows:
        return pd.DataFrame(columns=["Priority", "Question / Section", "Area", "What to check", "Why flagged", "Action requested from vessel/office", "HVPQ value", "PIQ value", "Class Status value", "Q88 value"])
    df = pd.DataFrame(rows)
    return df.drop_duplicates(subset=["Priority", "Question / Section", "Area", "What to check"])


def make_office_report_df(valid_df: pd.DataFrame, issues_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    rows.append({"Section": "A. Items appearing in order", "Status/Priority": "", "Summary / Issue": "", "Action": ""})
    if valid_df.empty:
        rows.append({"Section": "A. Items appearing in order", "Status/Priority": "No positive checks extracted", "Summary / Issue": "Review extracted fields and vessel checklist.", "Action": ""})
    else:
        for _, r in valid_df.iterrows():
            rows.append({"Section": r["Section"], "Status/Priority": r["Status"], "Summary / Issue": r["Office summary"], "Action": ""})
    rows.append({"Section": "B. Main issues / exceptions", "Status/Priority": "", "Summary / Issue": "Details and vessel questions are in the Vessel Action Checklist tab.", "Action": ""})
    if issues_df.empty:
        rows.append({"Section": "B. Main issues / exceptions", "Status/Priority": "No critical/high issue detected", "Summary / Issue": "Only manual targeted checks may remain for vessel confirmation.", "Action": ""})
    else:
        for _, r in issues_df.iterrows():
            rows.append({"Section": r["Area"], "Status/Priority": r["Priority"], "Summary / Issue": r["Issue"] + " — " + r["Brief office note"], "Action": r["Action"]})
    return pd.DataFrame(rows, columns=["Section", "Status/Priority", "Summary / Issue", "Action"])



# ----------------------------- v12 confidence, coverage and export helpers -----------------------------

def qno_for_finding(f: Finding) -> str:
    check_norm = normalize_key((f.check or "") + " " + (f.area or ""))
    if "cofr" in check_norm or "certificate" in check_norm or f.area == "Certificates": return "2.1.5"
    if "technical superintendent" in check_norm: return "PIQ 2.2.1001"
    if "marine superintendent" in check_norm: return "PIQ 2.2.1002"
    if "iws" in check_norm: return "1.5.5"
    if "special survey" in check_norm: return "1.5.6"
    if "dry dock" in check_norm or "drydock" in check_norm: return "1.5.4"
    if "annual survey" in check_norm: return "1.5.11"
    if "intermediate" in check_norm: return "1.5.12"
    if "cii" in check_norm or "aer" in check_norm: return "1.2.3"
    if "eexi" in check_norm: return "1.2.2"
    if "p&i" in check_norm or "pni" in check_norm: return "1.1.13"
    if "owner" in check_norm: return "1.3.1"
    if "type" in check_norm: return "1.1.8 / 1.1.9"
    if "incident" in check_norm: return "1.9.3 / 1.9.5 / PIQ 5.7"
    if "psc" in check_norm: return "1.9.8 / PIQ 2.8.2"
    if "tank" in check_norm: return "PIQ 2.3.3001-3003 / HVPQ 7.1.1 / 7.1.3"
    if "retrofit" in check_norm or "moc" in check_norm: return "PIQ 2.5.1002-1004"
    if "conditions of class" in check_norm: return "1.5.14"
    if "memoranda" in check_norm or "memo" in check_norm: return "1.5.16"
    return ""


def raw_text_for_source(page_cache: Dict[str, List[Tuple[int, str]]], source: str) -> str:
    return "\n".join(t for _, t in page_cache.get(source, []))


def hvpq_qid_status_df(obs_df: pd.DataFrame, hvpq_text: str) -> pd.DataFrame:
    base_cols = ["Priority", "Question No.", "Topic", "HVPQ check status", "What vessel/office should check", "HVPQ evidence excerpt", "Observation basis / example"]
    obs_pin = observation_pinpoint_rows(obs_df)
    if obs_pin.empty:
        return pd.DataFrame(columns=base_cols)
    rows = []
    text = hvpq_text or ""
    low = text.lower()
    for _, r in obs_pin.iterrows():
        q = str(r.get("Question No.", "")).strip()
        topic = str(r.get("Topic", "")).strip()
        action = str(r.get("What vessel should check", "")).strip()
        basis = str(r.get("Observation basis / example", "")).strip()
        status = "Could not reliably check - exact question number not found in HVPQ extraction"
        excerpt = ""
        if q:
            m = re.search(r"(?<!\d)" + re.escape(q) + r"(?!\d)", text)
            if m:
                start = max(0, m.start() - 80)
                end = min(len(text), m.start() + 900)
                snippet = clean_text(text[start:end])
                # cut before next major qid after the current one, if visible
                rel = snippet.find(q)
                if rel >= 0:
                    after = snippet[rel+len(q):]
                    nxt = re.search(r"\b\d{1,2}\.\d{1,2}(?:\.\d{1,4})?\b", after)
                    if nxt and nxt.start() > 80:
                        snippet = snippet[:rel+len(q)+nxt.start()]
                excerpt = snippet[:700]
                # basic blank/answer heuristic
                local = excerpt.lower()
                answer_tokens = [" yes", " no", "not applicable", " n/a", " na", " date", " issued", " expires", "expiry", "good", "class", "annual", "months", "202", "201", "bar", "mt", "usd"]
                if any(tok in local for tok in answer_tokens) and len(excerpt) > 80:
                    status = "Found in HVPQ extraction - value/answer appears present; verify exact correctness"
                else:
                    status = "Found in HVPQ extraction but answer could not be reliably confirmed; manual check required"
        rows.append({
            "Priority": "Observation-led check",
            "Question No.": q,
            "Topic": topic,
            "HVPQ check status": status,
            "What vessel/office should check": action,
            "HVPQ evidence excerpt": excerpt,
            "Observation basis / example": basis,
        })
    df = pd.DataFrame(rows, columns=base_cols)
    return df.drop_duplicates(subset=["Question No.", "Topic", "What vessel/office should check"])


def field_exists(fields: List[FieldRecord], field_id: str, sources: Optional[List[str]] = None) -> bool:
    for f in fields:
        if f.field_id == field_id and str(f.value).strip():
            if sources is None or f.source in sources:
                return True
    return False


def count_fields(fields: List[FieldRecord], prefix: str, sources: Optional[List[str]] = None) -> int:
    return sum(1 for f in fields if f.field_id.startswith(prefix) and str(f.value).strip() and (sources is None or f.source in sources))


def build_coverage_matrix(fields: List[FieldRecord], findings: List[Finding], obs_df: pd.DataFrame, hvpq_text: str) -> pd.DataFrame:
    rows = []
    def add(area, check, status, basis, manual_action=""):
        rows.append({"Area": area, "Check performed": check, "Status": status, "Basis / extracted evidence": basis, "Manual action if not reliable": manual_action})
    # Core document extraction
    add("Source extraction", "HVPQ identity and core fields", "Checked" if field_exists(fields,"vessel.imo",["HVPQ","XML"]) else "Could not reliably check", first_field(fields,"HVPQ","vessel.name") or first_field(fields,"XML","vessel.name") or "", "Confirm correct HVPQ was uploaded and PDF/XML extraction succeeded.")
    add("Source extraction", "PIQ fields", "Checked" if field_exists(fields,"vessel.name",["PIQ"]) else "Could not reliably check", first_field(fields,"PIQ","vessel.name"), "Upload PIQ PDF or check parser output.")
    add("Source extraction", "Class Status certificate/survey fields", "Checked" if count_fields(fields,"cert.",["CLASS"]) >= 3 or field_exists(fields,"classification.class_society",["CLASS"]) else "Could not reliably check", f"Class cert fields extracted: {count_fields(fields,'cert.',['CLASS'])}", "Class Status is authority for certificate/survey dates; verify manually if parser did not extract rows.")
    add("Source extraction", "Q88 value-add fields", "Checked" if field_exists(fields,"vessel.imo",["Q88"]) or count_fields(fields,"cert.",["Q88"]) >= 3 else "Not uploaded / could not reliably check", f"Q88 cert fields extracted: {count_fields(fields,'cert.',['Q88'])}", "Q88 is value-add only. Upload Q88 if available; mismatches should be highlighted separately.")
    # HVPQ as primary correction target
    add("HVPQ correction focus", "Certificate dates in HVPQ compared against Class Status where available", "Checked" if any(f.area=="Certificates" for f in findings) or count_fields(fields,"cert.",["HVPQ"]) else "Could not reliably check", f"HVPQ cert fields: {count_fields(fields,'cert.',['HVPQ'])}; Class cert fields: {count_fields(fields,'cert.',['CLASS'])}", "If Class Status rows are not parsed, check certificate table manually against HVPQ 2.1.5.")
    add("HVPQ correction focus", "Conditions/Memoranda/Dispensation declarations", "Checked" if any(field_exists(fields,fid,["HVPQ","CLASS","Q88"]) for fid in ["classification.conditions_of_class","classification.memo_of_class","classification.flag_dispensation"]) else "Could not reliably check", "Compared HVPQ/Q88/Class Status where extracted.", "Manually check Class Status for open conditions, memoranda, recommendations, exemptions, dispensations.")
    # PIQ-focused checks
    add("PIQ consistency", "Technical/Marine superintendent intervals", "Checked" if any(f.field_id.startswith("superintendent.") for f in fields) else "Could not reliably check", "Uses locked TS 7-month and MS 12-month strict rules.", "Check PIQ 2.2.1001 / 2.2.1002 visit rows manually.")
    add("PIQ consistency", "Tank inspection cycle", "Checked" if any(f.field_id.startswith("tank.") for f in fields) else "Could not reliably check", "Tank dates within cycle are not flagged as findings.", "Check PIQ 2.3.3001-3003 manually if dates not extracted.")
    add("PIQ/HVPQ/Q88", "Incident declarations / no-incident positive confirmation", "Checked" if any(f.field_id.startswith("incident.") for f in fields) else "Could not reliably check", "If all declare no incident, a positive vessel confirmation row is created.", "Confirm no reportable injury, machinery, mooring, navigation, pollution, security or operational incident in last 12 months.")
    # Observation library
    obs_q = hvpq_qid_status_df(obs_df, hvpq_text)
    if obs_df is None or obs_df.empty:
        add("Observation library", "Historical observation question-number checks", "Not uploaded", "No observation Excel uploaded.", "Upload HVPQ observation Excel to generate targeted HVPQ question checks.")
    elif obs_q.empty:
        add("Observation library", "Historical observation question-number checks", "Could not reliably check", "No exact question numbers captured from observation library.", "Review observation Excel format or add question numbers such as 10.1.4 in observation text.")
    else:
        missing = int(obs_q["HVPQ check status"].astype(str).str.contains("not found|could not reliably", case=False, regex=True).sum())
        add("Observation library", "Historical observation question-number checks", "Checked with manual gaps" if missing else "Checked", f"Generated {len(obs_q)} targeted HVPQ question checks; {missing} require manual confirmation.", "Review the Observation QID Checks tab for exact HVPQ question-level checks.")
    return pd.DataFrame(rows)


def make_hvpq_correction_register(findings: List[Finding]) -> pd.DataFrame:
    rows = []
    for f in findings:
        # Primary correction register: HVPQ vs authoritative/PIQ mismatches and high/medium manual checks.
        if f.risk.upper() not in ["CRITICAL", "HIGH", "MEDIUM"]:
            continue
        source_logic = ""
        action = f.action
        if f.class_value:
            source_logic = "Class Status is authoritative for certificate/class/survey data"
            if "update" not in action.lower() and f.hvpq_value:
                action = "If Class Status is current, correct HVPQ accordingly. " + action
        elif f.piq_value:
            source_logic = "PIQ cross-check against HVPQ / operational declaration"
        elif f.q88_value:
            source_logic = "Q88 value-add cross-check; not automatically authoritative"
        else:
            source_logic = "Manual verification required"
        rows.append({
            "Priority": f.risk,
            "HVPQ / PIQ question or section": qno_for_finding(f),
            "Area": f.area,
            "Issue / check": f.check,
            "Status": f.status,
            "Source logic": source_logic,
            "HVPQ value to review/correct": f.hvpq_value,
            "Class Status value (authoritative where applicable)": f.class_value,
            "PIQ value": f.piq_value,
            "Q88 value-add value": f.q88_value,
            "Reason": f.reason,
            "Required action": action,
        })
    cols = ["Priority","HVPQ / PIQ question or section","Area","Issue / check","Status","Source logic","HVPQ value to review/correct","Class Status value (authoritative where applicable)","PIQ value","Q88 value-add value","Reason","Required action"]
    return pd.DataFrame(rows, columns=cols)


def make_q88_value_add(findings: List[Finding]) -> pd.DataFrame:
    rows = []
    for f in findings:
        if f.q88_value and (not f.class_value or not f.hvpq_value or not semantically_equivalent(f.hvpq_value, f.q88_value, f.check)):
            rows.append({
                "Priority": f.risk,
                "Question / Section": qno_for_finding(f),
                "Area": f.area,
                "Q88 value-add issue": f.check,
                "HVPQ value": f.hvpq_value,
                "Q88 value": f.q88_value,
                "Class Status value": f.class_value,
                "Interpretation": "Q88 is a value-add cross-check. Do not blindly change HVPQ based on Q88 alone unless Class Status/certificate/source evidence supports it.",
                "Action": f.action,
            })
    return pd.DataFrame(rows, columns=["Priority","Question / Section","Area","Q88 value-add issue","HVPQ value","Q88 value","Class Status value","Interpretation","Action"])


def make_manual_unchecked(coverage_df: pd.DataFrame, obs_qid_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if not coverage_df.empty:
        bad = coverage_df[coverage_df["Status"].astype(str).str.contains("Could not reliably|Not uploaded", case=False, regex=True, na=False)]
        for _, r in bad.iterrows():
            rows.append({"Priority": "Manual verification", "Area": r["Area"], "Item not reliably checked": r["Check performed"], "Why": r["Basis / extracted evidence"], "What to do": r["Manual action if not reliable"]})
    if not obs_qid_df.empty:
        bad2 = obs_qid_df[obs_qid_df["HVPQ check status"].astype(str).str.contains("could not reliably|not found|manual", case=False, regex=True, na=False)]
        for _, r in bad2.iterrows():
            rows.append({"Priority": "Observation-led manual check", "Area": r.get("Topic",""), "Item not reliably checked": f"HVPQ {r.get('Question No.','')}", "Why": r.get("HVPQ check status",""), "What to do": r.get("What vessel/office should check","")})
    return pd.DataFrame(rows, columns=["Priority","Area","Item not reliably checked","Why","What to do"])


def make_vessel_action_checklist(findings: List[Finding], obs_df: pd.DataFrame, hvpq_text: str = "") -> pd.DataFrame:
    rows = []
    for f in findings:
        if f.risk.upper() in ["CRITICAL", "HIGH", "MEDIUM"]:
            rows.append({
                "Priority": f.risk,
                "Question / Section": qno_for_finding(f),
                "Area": f.area,
                "What vessel/office should check": f.check,
                "Why it is being asked": f.reason,
                "Action requested": f.action,
                "HVPQ value": f.hvpq_value,
                "PIQ value": f.piq_value,
                "Class Status value": f.class_value,
                "Q88 value": f.q88_value,
            })
    obs_qid = hvpq_qid_status_df(obs_df, hvpq_text)
    for _, r in obs_qid.iterrows():
        rows.append({
            "Priority": "Targeted check",
            "Question / Section": r.get("Question No.", ""),
            "Area": r.get("Topic", "Observation library"),
            "What vessel/office should check": r.get("What vessel/office should check", ""),
            "Why it is being asked": f"Historical observation pattern. HVPQ extraction status: {r.get('HVPQ check status','')}",
            "Action requested": "Verify the HVPQ answer and supporting evidence; correct HVPQ if wrong/blank/stale.",
            "HVPQ value": r.get("HVPQ evidence excerpt", ""),
            "PIQ value": "", "Class Status value": "", "Q88 value": "",
        })
    if not rows:
        return pd.DataFrame(columns=["Priority", "Question / Section", "Area", "What vessel/office should check", "Why it is being asked", "Action requested", "HVPQ value", "PIQ value", "Class Status value", "Q88 value"])
    df = pd.DataFrame(rows)
    return df.drop_duplicates(subset=["Priority", "Question / Section", "Area", "What vessel/office should check"])


def make_excel(findings: List[Finding], fields: List[FieldRecord], vessel_actions: pd.DataFrame, obs_qids: pd.DataFrame, coverage_df: pd.DataFrame, manual_df: pd.DataFrame, q88_df: pd.DataFrame, hvpq_register: pd.DataFrame) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        vessel_actions.to_excel(writer, index=False, sheet_name="Vessel Register")
        hvpq_register.to_excel(writer, index=False, sheet_name="HVPQ PIQ Issues")
        manual_df.to_excel(writer, index=False, sheet_name="Manual Confirmation")
        q88_df.to_excel(writer, index=False, sheet_name="Q88 Value Add")
        coverage_df.to_excel(writer, index=False, sheet_name="Coverage Matrix")
        obs_qids.to_excel(writer, index=False, sheet_name="Observation Q Checks")
        df_from_findings(findings).to_excel(writer, index=False, sheet_name="Detailed Findings")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            # Header style
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
            widths = {
                "A": 18, "B": 22, "C": 24, "D": 34, "E": 22, "F": 44, "G": 44, "H": 44, "I": 44, "J": 44, "K": 52, "L": 52
            }
            for col_idx, col in enumerate(ws.columns, start=1):
                letter = get_column_letter(col_idx)
                max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in col)
                ws.column_dimensions[letter].width = min(max(widths.get(letter, 12), min(max_len + 2, 55)), 60)
            # readable row heights for wrapped text
            for r in range(2, min(ws.max_row, 250) + 1):
                ws.row_dimensions[r].height = 48
    return bio.getvalue()


# ----------------------------- Simple UI summary helpers -----------------------------

def _risk_n(findings: List[Finding], risk: str) -> int:
    return sum(1 for f in findings if f.risk.upper() == risk.upper())


def _area_has_issue(findings: List[Finding], area_keyword: str, risks=("CRITICAL", "HIGH", "MEDIUM")) -> bool:
    k = area_keyword.lower()
    return any(k in f.area.lower() and f.risk.upper() in risks for f in findings)


def build_human_review_summary(fields: List[FieldRecord], findings: List[Finding], coverage_df: pd.DataFrame, manual_df: pd.DataFrame, obs_qid_df: pd.DataFrame, ref_date: date) -> Dict[str, str]:
    checked = []
    if field_exists(fields, "vessel.imo", ["HVPQ", "XML"]): checked.append("HVPQ identity/core particulars")
    if count_fields(fields, "cert.", ["HVPQ"]): checked.append("HVPQ certificate table")
    if count_fields(fields, "cert.", ["CLASS"]): checked.append("Class Status certificate/survey dates")
    if field_exists(fields, "vessel.imo", ["Q88"]) or count_fields(fields, "cert.", ["Q88"]): checked.append("Q88 value-add fields")
    if any(f.field_id.startswith("superintendent.") for f in fields): checked.append("PIQ superintendent visit gaps")
    if any(f.field_id.startswith("tank.") for f in fields): checked.append("PIQ tank inspection cycles")
    if any(f.field_id.startswith("incidents.") for f in fields): checked.append("incident/no-incident declarations")
    if not obs_qid_df.empty: checked.append("historical observation question numbers")
    checked_text = ", ".join(checked) if checked else "the uploaded documents, subject to extraction quality"

    ok_bits = []
    # Tank inspections ok if no high tank issue and PIQ dates exist within cycle
    tank_ok = []
    for title, old_fid, freq_fid in [("cargo/slop", "tank.cargo_slop.oldest", "tank.cargo_slop.freq_months"), ("ballast", "tank.ballast.oldest", "tank.ballast.freq_months"), ("void", "tank.void.oldest", "tank.void.freq_months")]:
        old = parse_date_any(first_field(fields, "PIQ", old_fid))
        freq = first_field(fields, "PIQ", freq_fid)
        months = 12
        try: months = int(float(freq)) if freq else 12
        except Exception: pass
        if old and old + relativedelta(months=months) >= ref_date:
            tank_ok.append(title)
    if tank_ok: ok_bits.append(f"tank inspection cycles appear within required interval for {', '.join(tank_ok)} tanks")
    if not _area_has_issue(findings, "Management Oversight", ("CRITICAL", "HIGH")) and any(f.field_id.startswith("superintendent.") for f in fields):
        ok_bits.append("no high-risk superintendent interval breach was detected from extracted PIQ dates")
    if not _area_has_issue(findings, "Certificates", ("CRITICAL", "HIGH")) and count_fields(fields, "cert.", ["HVPQ"]):
        ok_bits.append("no high-risk certificate mismatch/expiry issue was detected from extracted certificate rows")
    if not _area_has_issue(findings, "Class", ("CRITICAL", "HIGH")) and any(field_exists(fields, fid, ["HVPQ", "CLASS", "Q88"]) for fid in ["classification.conditions_of_class", "classification.memo_of_class"]):
        ok_bits.append("Class Conditions/Memoranda were checked where extractable")
    ok_text = "; ".join(ok_bits) if ok_bits else "no clean 'in-order' conclusion is shown where extraction was insufficient; those items are moved to manual confirmation instead"

    main_issues = [f for f in findings if f.risk.upper() in ["CRITICAL", "HIGH"]]
    if main_issues:
        issue_text = "; ".join([f"{f.area}: {f.check}" for f in main_issues[:8]])
        if len(main_issues) > 8: issue_text += f"; plus {len(main_issues)-8} more high-priority item(s)"
    else:
        issue_text = "no critical/high-priority mismatch was detected from the extracted mapped fields"

    if not manual_df.empty:
        manual_text = f"{len(manual_df)} item(s) could not be reliably checked or need positive confirmation. These are included in the Manual Confirmation and Vessel Action Checklist tabs so they are not missed."
    else:
        manual_text = "no major extraction-gap/manual-confirmation item was generated from the current upload."

    if obs_qid_df.empty:
        obs_text = "No observation-library question-number checks were generated. Upload an observation Excel containing HVPQ question numbers such as 10.1.4 or 2.1.5 to activate this section."
    else:
        review = obs_qid_df[obs_qid_df["HVPQ check status"].astype(str).str.contains("could not reliably|not found|manual", case=False, regex=True, na=False)]
        ok = len(obs_qid_df) - len(review)
        obs_text = f"{len(obs_qid_df)} repeated-observation HVPQ question check(s) were generated from the observation sheet. {ok} were located with an apparent answer/excerpt; {len(review)} need manual review or clearer evidence."

    return {
        "checked": f"Checked: {checked_text}.",
        "ok": f"Appearing in order: {ok_text}.",
        "issues": f"Not satisfactory / requires correction: {issue_text}.",
        "manual": f"Could not reliably check / manual confirmation: {manual_text}",
        "obs": f"Repeat-observation coverage: {obs_text}",
    }


def build_major_repeat_summary(obs_qid_df: pd.DataFrame) -> pd.DataFrame:
    cols = ["Question No.", "Topic", "Review status", "What to check", "HVPQ evidence excerpt"]
    if obs_qid_df is None or obs_qid_df.empty:
        return pd.DataFrame(columns=cols)
    rows = []
    for _, r in obs_qid_df.iterrows():
        stt = str(r.get("HVPQ check status", ""))
        if re.search(r"could not reliably|not found|manual", stt, flags=re.I):
            status = "Needs review"
        else:
            status = "Located in HVPQ - verify correctness"
        rows.append({
            "Question No.": r.get("Question No.", ""),
            "Topic": r.get("Topic", ""),
            "Review status": status,
            "What to check": r.get("What vessel/office should check", ""),
            "HVPQ evidence excerpt": r.get("HVPQ evidence excerpt", ""),
        })
    return pd.DataFrame(rows, columns=cols)


def style_priority_dataframe(df: pd.DataFrame):
    if df is None or df.empty:
        return df
    def row_style(row):
        val = " ".join(str(x).upper() for x in row.values)
        if "CRITICAL" in val:
            return ['background-color: #fde2e1'] * len(row)
        if "HIGH" in val:
            return ['background-color: #fff1cc'] * len(row)
        if "NEEDS REVIEW" in val or "MANUAL" in val or "COULD NOT" in val:
            return ['background-color: #eef4ff'] * len(row)
        return [''] * len(row)
    return df.style.apply(row_style, axis=1)

# ----------------------------- Streamlit app -----------------------------

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.caption(APP_SUBTITLE)

    with st.sidebar:
        st.header("Upload documents")
        hvpq_file = st.file_uploader("HVPQ PDF — main document to correct", type=["pdf"], key="hvpq")
        hvpq_xml = st.file_uploader("HVPQ XML (optional)", type=["xml"], key="xml")
        piq_file = st.file_uploader("PIQ PDF", type=["pdf"], key="piq")
        q88_file = st.file_uploader("Q88 PDF — value-add cross-check", type=["pdf"], key="q88")
        class_file = st.file_uploader("Class Status PDF — certificate/survey authority", type=["pdf"], key="class")
        obs_file = st.file_uploader("HVPQ observation library Excel", type=["xlsx", "xls"], key="obs")
        inc_obs_file = st.file_uploader("Incident observation library Excel", type=["xlsx", "xls"], key="incobs")
        st.divider()
        ref_date_input = st.date_input("Reference / review date", value=date.today())
        show_low = st.checkbox("Show low-risk findings", value=False)
        use_llm = st.checkbox("Use local LLM extraction assist (Ollama)", value=False)
        ollama_url = st.text_input("Ollama URL", value="http://localhost:11434", disabled=not use_llm)
        ollama_model = st.text_input("Ollama model", value="qwen2.5:14b", disabled=not use_llm)
        run_btn = st.button("Run checks", type="primary")

    if not run_btn:
        st.info("Upload HVPQ, PIQ, Class Status and Q88 where available, then click **Run checks**. HVPQ is treated as the correction target; Class Status is used only for certificate/survey dates and Conditions/Memoranda; Q88 is shown separately as value-add.")
        return

    settings = {"show_low": show_low}
    all_fields: List[FieldRecord] = []
    page_cache = {}

    with st.spinner("Extracting mapped fields and running verification rules..."):
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

    # Keep the extracted table out of the main workflow. It is available only in Advanced Review.
    edited_fields = all_fields

    findings = run_rules(edited_fields, ref_date_input, settings, obs_df)
    if not show_low:
        findings = [f for f in findings if f.risk.upper() != "LOW"]
    findings_df = df_from_findings(findings)
    hvpq_text = raw_text_for_source(page_cache, "HVPQ")
    obs_qid_df = hvpq_qid_status_df(obs_df, hvpq_text)
    coverage_df = build_coverage_matrix(edited_fields, findings, obs_df, hvpq_text)
    manual_df = make_manual_unchecked(coverage_df, obs_qid_df)
    hvpq_register_df = make_hvpq_correction_register(findings)
    q88_value_add_df = make_q88_value_add(findings)
    vessel_actions_df = make_vessel_action_checklist(findings, obs_df, hvpq_text)
    repeat_summary_df = build_major_repeat_summary(obs_qid_df)
    summary = build_human_review_summary(edited_fields, findings, coverage_df, manual_df, obs_qid_df, ref_date_input)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Critical / High", _risk_n(findings, "CRITICAL") + _risk_n(findings, "HIGH"))
    c2.metric("HVPQ correction rows", len(hvpq_register_df))
    c3.metric("Manual confirmations", len(manual_df))
    c4.metric("Observation Q checks", len(obs_qid_df))

    tabs = st.tabs(["Review Dashboard", "Vessel Register", "HVPQ / PIQ Issues", "Manual Confirmation", "Q88 Value Add", "Coverage", "Advanced Review"])

    with tabs[0]:
        st.subheader("Review dashboard")
        st.markdown("""
        This dashboard is meant to give the reviewer confidence that the uploaded documents were checked in a controlled way.  
        **HVPQ is the main document to correct. Class Status is used as the authority only for certificate/survey dates and Conditions/Memoranda. Q88 is treated as a value-add cross-check and is shown separately.**
        """)
        st.success(summary["checked"] + "\n\n" + summary["ok"])
        if findings_df.empty or not any(r in findings_df.get("risk", pd.Series(dtype=str)).astype(str).str.upper().tolist() for r in ["CRITICAL", "HIGH"]):
            st.info(summary["issues"])
        else:
            st.error(summary["issues"])
        if manual_df.empty:
            st.success(summary["manual"])
        else:
            st.warning(summary["manual"] + " These items are included in the Vessel Register / Manual Confirmation tabs.")
        st.info(summary["obs"])

        st.markdown("### Major repeat-finding checks from observation sheet")
        if repeat_summary_df.empty:
            st.caption("No repeat-finding question-number checks generated. Upload observation sheets with HVPQ question numbers to activate this section.")
        else:
            st.dataframe(style_priority_dataframe(repeat_summary_df), use_container_width=True, height=360)

        st.markdown("### Top action items")
        top_actions = vessel_actions_df.head(12) if not vessel_actions_df.empty else pd.DataFrame()
        if top_actions.empty:
            st.success("No action item generated from the mapped checks. Review Manual Confirmation if any document was not uploaded or not reliably extracted.")
        else:
            st.dataframe(style_priority_dataframe(top_actions), use_container_width=True, height=420)

    with tabs[1]:
        st.subheader("Vessel register — clear action list to send to vessel/office")
        st.caption("This combines confirmed mismatches, blank/not-extracted mapped entries, no-incident confirmation, and observation-led targeted checks. Class Status values are shown only as reference where relevant.")
        st.dataframe(style_priority_dataframe(vessel_actions_df), use_container_width=True, height=620)

    with tabs[2]:
        st.subheader("HVPQ / PIQ issues and correction register")
        st.caption("Use this as the office correction register. HVPQ is the correction target. PIQ mismatches/blanks are shown where operational declarations need alignment.")
        st.dataframe(style_priority_dataframe(hvpq_register_df), use_container_width=True, height=620)

    with tabs[3]:
        st.subheader("Manual confirmation / could not reliably check")
        st.caption("These rows are deliberately shown so the reviewer knows what the app could not verify with confidence. They should be checked manually before closing the review.")
        if manual_df.empty:
            st.success("No manual-confirmation row generated from the current upload.")
        else:
            st.dataframe(style_priority_dataframe(manual_df), use_container_width=True, height=620)

    with tabs[4]:
        st.subheader("Q88 value-add mismatches / blanks")
        st.caption("Q88 is not the authority. Use this tab to spot Q88/HVPQ inconsistencies, blanks and stale entries, then verify against source evidence before correcting HVPQ.")
        if q88_value_add_df.empty:
            st.success("No Q88 value-add mismatch detected from mapped fields.")
        else:
            st.dataframe(style_priority_dataframe(q88_value_add_df), use_container_width=True, height=560)

    with tabs[5]:
        st.subheader("Coverage — what was checked vs what needs manual confirmation")
        st.caption("This is the confidence matrix. Items marked 'Could not reliably check' are also carried into Manual Confirmation/Vessel Register where applicable.")
        st.dataframe(style_priority_dataframe(coverage_df), use_container_width=True, height=560)
        st.markdown("### Observation-library HVPQ question checks")
        if obs_qid_df.empty:
            st.info("No exact HVPQ question-number checks generated from the observation Excel.")
        else:
            st.dataframe(style_priority_dataframe(obs_qid_df), use_container_width=True, height=460)

    with tabs[6]:
        st.subheader("Advanced review / audit trail")
        st.caption("Kept out of the main workflow to avoid confusing users. Use only for troubleshooting extraction.")
        with st.expander("Show extracted structured fields"):
            st.dataframe(df_from_fields(edited_fields), use_container_width=True, height=500)
        with st.expander("Show raw page text"):
            src = st.selectbox("Source", list(page_cache.keys()) or ["None"])
            if src != "None":
                page_no = st.number_input("Page", min_value=1, max_value=max([p for p, _ in page_cache[src]]), value=1)
                txt = dict(page_cache[src]).get(page_no, "")
                st.text_area("Extracted page text", txt, height=500)
        with st.expander("Show JSON output"):
            st.json({"findings": [asdict(f) for f in findings], "fields": [asdict(f) for f in edited_fields], "coverage": coverage_df.to_dict(orient="records")})

    xlsx = make_excel(findings, edited_fields, vessel_actions_df, obs_qid_df, coverage_df, manual_df, q88_value_add_df, hvpq_register_df)
    st.download_button("Download Excel register", xlsx, file_name="hvpq_piq_q88_class_verification_register_v13.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
