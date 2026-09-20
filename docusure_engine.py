"""Evidence-quality, document classification and readiness helpers for DocuSure.

The Streamlit application deliberately keeps extraction and judgement separate:
this module never turns a missing extraction into a confirmed defect.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from datetime import date
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
from dateutil import parser as dateparser


DOC_LABELS = {
    "HVPQ": "HVPQ",
    "PIQ": "PIQ",
    "Q88": "Q88",
    "CLASS": "Class Status / Survey Status",
    "CERTIFICATE": "Statutory / trading certificate",
    "PNI": "P&I / insurance evidence",
    "STS_PLAN": "STS transfer plan / procedure",
    "STS_ASSESSMENT": "STS compatibility / risk assessment / JPO",
    "PSC_REPORT": "Port State Control report",
    "MOORING_PLAN": "Mooring line management / mooring plan",
    "CREW_MATRIX": "Crew matrix / manning document",
    "PORT_HISTORY": "Last 10 ports / port history",
    "SANCTIONS": "Sanctions / port screening evidence",
    "HVPQ_XML": "HVPQ XML",
    "UNKNOWN": "Unclassified document",
}


CERTIFICATE_NAMES = {
    "safety_equipment": "Cargo Ship Safety Equipment",
    "safety_radio": "Cargo Ship Safety Radio",
    "safety_construction": "Cargo Ship Safety Construction",
    "loadline": "International Load Line",
    "iopp": "International Oil Pollution Prevention",
    "bwm": "International Ballast Water Management",
    "smc": "Safety Management Certificate",
    "doc": "Document of Compliance",
    "issc": "International Ship Security Certificate",
    "uscg_coc": "USCG Certificate of Compliance",
    "clc_oil": "CLC Oil Pollution",
    "clc_bunker": "Bunker Oil Pollution",
    "wreck_removal": "Wreck Removal",
    "cofr": "Certificate of Financial Responsibility",
    "vgp": "Vessel General Permit",
    "cof_chemical": "Certificate of Fitness",
    "class_certificate": "Certificate of Class",
    "tonnage": "International Tonnage",
    "iapp": "International Air Pollution Prevention",
    "ispp": "International Sewage Pollution Prevention",
    "mlc": "Maritime Labour Certificate",
    "ship_sanitation": "Ship Sanitation Certificate",
    "pni_cover": "P&I cover",
}


SOURCE_PRIORITY = {
    "CERTIFICATE": 6,
    "CLASS": 5,
    "HVPQ": 4,
    "PIQ": 3,
    "Q88": 2,
    "XML": 1,
    "CREW_MATRIX": 3,
    "PORT_HISTORY": 3,
    "SANCTIONS": 3,
}


def clean(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\x00", " ")).strip()


def normalise(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()


def valid_imo(value: Any) -> bool:
    digits = re.sub(r"\D", "", clean(value))
    if len(digits) != 7:
        return False
    return sum(int(digits[i]) * (7 - i) for i in range(6)) % 10 == int(digits[-1])


MONTH_TOKEN = r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
DATE_TOKEN_RE = re.compile(
    rf"\b(?:\d{{1,2}}[\s./-]+{MONTH_TOKEN}[\s,./-]+\d{{2,4}}|"
    rf"{MONTH_TOKEN}[\s./-]+\d{{1,2}},?[\s./-]+\d{{2,4}}|"
    r"\d{4}-\d{1,2}-\d{1,2}|\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\b",
    re.I,
)


def parse_date(value: Any) -> Optional[date]:
    match = DATE_TOKEN_RE.search(clean(value))
    if not match:
        return None
    token = match.group(0)
    yearfirst = bool(re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", token))
    try:
        parsed = dateparser.parse(token, dayfirst=not yearfirst, yearfirst=yearfirst, fuzzy=False)
    except Exception:
        return None
    if parsed and 1900 <= parsed.year <= 2100:
        return parsed.date()
    return None


def dates_in_text(text: str) -> List[date]:
    result: List[date] = []
    for match in DATE_TOKEN_RE.finditer(text or ""):
        parsed = parse_date(match.group(0))
        if parsed:
            result.append(parsed)
    return result


@dataclass
class DocumentProfile:
    filename: str
    doc_type: str
    doc_label: str
    confidence_score: int
    confidence: str
    classification_evidence: str
    imo: str = ""
    vessel_name: str = ""
    pages: int = 0
    pages_with_text: int = 0
    text_chars: int = 0
    text_quality: str = ""
    assigned_group: str = ""
    related_imos: str = ""
    grouping_confidence: str = ""
    grouping_evidence: str = ""
    document_format: str = "pdf"
    machine_readable: str = "Yes"

    def row(self) -> Dict[str, Any]:
        return asdict(self)


TYPE_SIGNATURES: Dict[str, Sequence[Tuple[str, int, str]]] = {
    "HVPQ": (
        (r"date\s+(?:this\s+)?hvpq\s+document\s+completed", 12, "HVPQ completion field"),
        (r"harmoni[sz]ed\s+vessel\s+particulars", 12, "HVPQ title"),
        (r"provide\s+details\s+for\s+mooring\s+ropes", 4, "HVPQ mooring section"),
        (r"IMO/LR\s+Number", 3, "HVPQ identity format"),
    ),
    "PIQ": (
        (r"\bpre[- ]inspection\s+questionnaire\b", 12, "PIQ title"),
        (r"\bPIQ\s+Report\b", 12, "PIQ report title"),
        (r"technical\s+superintendent\s+inspection\s+completed", 6, "PIQ superintendent table"),
        (r"marine\s+superintendent\s+inspection\s+completed", 6, "PIQ superintendent table"),
    ),
    "Q88": (
        (r"\bQ88\b", 10, "Q88 identifier"),
        (r"date\s+(?:the\s+)?Q88\s+(?:was\s+)?updated", 10, "Q88 update field"),
        (r"Q88\.com", 8, "Q88 publisher"),
    ),
    "CLASS": (
        (r"\bclass\s+status\b|\bsurvey\s+status\b", 11, "Class/survey status title"),
        (r"certificate\s+description", 7, "Class certificate table"),
        (r"classification\s+status|status\s+of\s+surveys", 7, "Class status wording"),
        (r"conditions?\s+of\s+class.*memorand", 5, "Class conditions/memoranda sections"),
    ),
    "PNI": (
        (r"certificate\s+of\s+entry", 12, "Certificate of Entry"),
        (r"protection\s+(?:and|&)\s+indemnity|P\s*(?:and|&)\s*I\s+Club", 6, "P&I wording"),
        (r"blue\s+card", 5, "P&I blue card"),
    ),
    "STS_PLAN": (
        (r"ship\s+to\s+ship\s+(?:transfer\s+)?(?:operations?\s+)?plan", 14, "STS plan title"),
        (r"\bSTS\s+(?:operations?\s+)?plan\b", 12, "STS plan wording"),
        (r"MARPOL.*Annex\s+I.*(?:Regulation|Reg\.?)[ ]*41", 7, "MARPOL STS plan reference"),
    ),
    "STS_ASSESSMENT": (
        (r"joint\s+plan\s+of\s+operation|\bJPO\b", 12, "Joint plan of operation"),
        (r"STS\s+(?:compatibility|risk\s+assessment|clearance|questionnaire)", 11, "STS assessment title"),
        (r"ship\s+compatibility\s+study", 10, "Compatibility study"),
    ),
    "PSC_REPORT": (
        (r"port\s+state\s+control\s+(?:inspection\s+)?report", 13, "PSC report title"),
        (r"report\s+of\s+inspection\s+in\s+accordance\s+with\s+the\s+Paris\s+MOU", 10, "PSC report wording"),
        (r"deficienc(?:y|ies).*detention", 5, "PSC findings wording"),
    ),
    "MOORING_PLAN": (
        (r"line\s+management\s+plan|\bLMP\b", 12, "Line Management Plan"),
        (r"mooring\s+system\s+management\s+plan|\bMSMP\b", 12, "Mooring System Management Plan"),
        (r"mooring\s+rope\s+inspection\s+(?:and\s+)?retirement", 7, "Mooring management wording"),
    ),
    "CREW_MATRIX": (
        (r"crew\s+(?:matrix|list|complement)", 13, "Crew matrix/list title"),
        (r"minimum\s+safe\s+manning", 11, "Minimum safe manning wording"),
        (r"matrix\s+of\s+(?:competence|crew)", 9, "Crew competence matrix wording"),
        (r"\b(?:master|chief\s+officer|chief\s+engineer|second\s+engineer|2nd\s+engineer)\b", 4, "Marine rank table"),
    ),
    "PORT_HISTORY": (
        (r"last\s+(?:10|ten)\s+ports?", 14, "Last ten ports heading"),
        (r"port\s+history|previous\s+ports?|ports?\s+of\s+call", 10, "Port history wording"),
        (r"last\s+port.*(?:arrival|departure|sailing)", 7, "Last port movement wording"),
    ),
    "SANCTIONS": (
        (r"sanction(?:s)?\s+(?:screening|declaration|check|list)", 13, "Sanctions screening wording"),
        (r"screened\s+against\s+(?:sanctions|sanction\s+lists?)", 11, "Sanctions screening assertion"),
        (r"high\s+risk\s+ports?", 8, "High-risk port wording"),
    ),
    "CERTIFICATE": (
        (r"\bcertificate\s+(?:number|no\.?|of)\b", 5, "Certificate wording"),
        (r"\bvalid\s+(?:until|to)\b|\bexpiry\s+date\b", 5, "Certificate validity field"),
        (r"international\s+(?:oil|air|sewage|ballast|load\s+line|tonnage)", 5, "Statutory certificate title"),
        (r"cargo\s+ship\s+safety\s+(?:equipment|radio|construction)", 6, "SOLAS certificate title"),
    ),
}


FILENAME_HINTS = {
    "HVPQ": ("hvpq",),
    "PIQ": ("piq", "pre inspection questionnaire"),
    "Q88": ("q88",),
    "CLASS": ("class status", "survey status", "classstatus"),
    "PNI": ("p&i", "pni", "certificate of entry", "insurance"),
    "STS_PLAN": ("sts plan", "sts procedure"),
    "STS_ASSESSMENT": ("sts assessment", "compatibility", "jpo", "clearance"),
    "PSC_REPORT": ("psc", "port state"),
    "MOORING_PLAN": ("lmp", "msmp", "mooring plan", "line management"),
    "CREW_MATRIX": ("crew matrix", "crew list", "manning", "complement", "crew"),
    "PORT_HISTORY": ("last 10", "last10", "port history", "ports", "port calls"),
    "SANCTIONS": ("sanction", "screening", "high risk port"),
    "CERTIFICATE": ("certificate", "cert ", "coc", "iopp", "smc", "issc"),
}


def _confidence_label(score: int) -> str:
    if score >= 85:
        return "High"
    if score >= 65:
        return "Medium"
    return "Low"


def detect_document_type(filename: str, text: str, is_xml: bool = False) -> Tuple[str, int, str]:
    name = normalise(filename)
    body = text or ""
    if is_xml or filename.lower().endswith(".xml"):
        if re.search(r"<\s*(?:Vessel|Document|Template)\b", body, re.I):
            return "HVPQ_XML", 93, "XML vessel/document structure"

    scores: Dict[str, int] = {}
    evidence: Dict[str, List[str]] = {}
    for doc_type, signatures in TYPE_SIGNATURES.items():
        scores[doc_type] = 0
        evidence[doc_type] = []
        for pattern, weight, label in signatures:
            if re.search(pattern, body, re.I | re.S):
                scores[doc_type] += weight
                evidence[doc_type].append(label)
        for hint in FILENAME_HINTS.get(doc_type, ()):
            if hint in name:
                scores[doc_type] += 5
                evidence[doc_type].append("filename")

    ordered = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    best_type, best_score = ordered[0] if ordered else ("UNKNOWN", 0)
    second_score = ordered[1][1] if len(ordered) > 1 else 0
    if best_score < 7:
        return "UNKNOWN", min(49, best_score * 6), "No reliable document signature"
    # Convert rule points into a transparent confidence score and reduce close calls.
    confidence = min(98, 55 + best_score * 3)
    if best_score - second_score <= 2:
        confidence = min(confidence, 64)
    return best_type, confidence, ", ".join(dict.fromkeys(evidence[best_type]))


def detect_imo(text: str) -> Tuple[str, int, str]:
    candidates: List[Tuple[int, str, str]] = []
    for match in re.finditer(r"\b\d{7}\b", text or ""):
        value = match.group(0)
        if not valid_imo(value):
            continue
        context = clean((text or "")[max(0, match.start() - 55):match.end() + 30])
        score = 82
        if re.search(r"\bIMO(?:/LR)?(?:\s+(?:No\.?|Number))?\b", context, re.I):
            score = 98
        candidates.append((score, value, context))
    if not candidates:
        return "", 0, ""
    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][0], candidates[0][2]


def detect_all_imos(text: str) -> List[str]:
    """Return checksum-valid IMO numbers in first-visible order."""
    values: List[str] = []
    for match in re.finditer(r"\b\d{7}\b", text or ""):
        value = match.group(0)
        if valid_imo(value) and value not in values:
            values.append(value)
    return values


def detect_vessel_name(text: str, filename: str = "") -> str:
    patterns = (
        r"(?:Name\s+of\s+(?:the\s+)?(?:Ship|Vessel)|Ship\s+Name|Vessel\s+Name)\s*[:\-]?\s*\n?\s*([A-Z0-9][A-Z0-9 '\-/]{2,60})",
        r"IMO/LR\s+Number\s+\d{7}\s+([A-Z0-9][A-Z0-9 '\-/]{2,60})",
        r"\bPIQ\s+Report\s+([A-Z0-9][A-Z0-9 '\-/]{2,60})",
    )
    for pattern in patterns:
        match = re.search(pattern, text or "", re.I)
        if match:
            value = clean(match.group(1)).split(" Date ")[0]
            value = re.split(r"\s{2,}|\n", value)[0]
            if 2 < len(value) <= 60 and not re.search(r"questionnaire|report|number", value, re.I):
                return value
    base = re.sub(r"\.[A-Za-z0-9]{1,5}$", "", filename)
    base = re.sub(r"(?i)\b(hvpq|piq|q88|class status|certificate|survey status|sts|report)\b", " ", base)
    base = clean(re.sub(r"[_-]+", " ", base))
    return base[:60] if 2 < len(base) <= 60 else ""


def profile_document(filename: str, pages: Sequence[Tuple[int, str]], text: str, is_xml: bool = False, document_format: str = "") -> DocumentProfile:
    doc_type, type_score, evidence = detect_document_type(filename, text, is_xml=is_xml)
    imo, _, _ = detect_imo(text)
    related_imos = detect_all_imos(text)
    page_count = len(pages) if pages else (1 if text else 0)
    text_pages = sum(1 for _, page_text in pages if len(re.sub(r"\s+", "", page_text or "")) >= 50)
    chars = len(re.sub(r"\s+", "", text or ""))
    ratio = text_pages / page_count if page_count else 0.0
    if chars < 100 or ratio < 0.25:
        quality = "Poor — scanned/unreadable or OCR unavailable"
    elif ratio < 0.75:
        quality = "Partial — some pages have little searchable text"
    else:
        quality = "Good — searchable text detected"
    return DocumentProfile(
        filename=filename,
        doc_type=doc_type,
        doc_label=DOC_LABELS.get(doc_type, doc_type),
        confidence_score=type_score,
        confidence=_confidence_label(type_score),
        classification_evidence=evidence,
        imo=imo,
        vessel_name=detect_vessel_name(text, filename),
        pages=page_count,
        pages_with_text=text_pages,
        text_chars=chars,
        text_quality=quality,
        assigned_group=imo or "Unassigned",
        related_imos=", ".join(related_imos),
        grouping_confidence="High" if imo and len(related_imos) == 1 else ("Medium" if imo else "Unassigned"),
        grouping_evidence=("Checksum-valid IMO extracted" if imo and len(related_imos) == 1 else (f"Multiple checksum-valid IMOs visible: {', '.join(related_imos)}" if related_imos else "No checksum-valid IMO extracted")),
        document_format=document_format or ("xml" if is_xml else "pdf"),
        machine_readable="Yes" if chars >= 100 else "Partial",
    )


def document_inventory_df(profiles: Sequence[DocumentProfile]) -> pd.DataFrame:
    columns = [
        "File", "Format", "Detected document", "Assigned vessel group", "Vessel / IMO",
        "Group confidence", "Why grouped", "Classification confidence",
        "Why classified", "Pages", "Text quality", "Machine-readable", "Action",
    ]
    rows = []
    for profile in profiles:
        vessel = " / ".join(x for x in [profile.vessel_name, profile.imo] if x) or "Not identified"
        related = [value.strip() for value in profile.related_imos.split(",") if value.strip()]
        assigned_display = f"Shared STS: {', '.join(related)}" if profile.doc_type == "STS_ASSESSMENT" and len(related) > 1 else (profile.assigned_group or "Unassigned")
        action = "Ready"
        if profile.doc_type == "UNKNOWN":
            action = "Open and identify manually; excluded from deterministic checks"
        elif profile.confidence_score < 65:
            action = "Confirm classification before relying on results"
        if profile.text_quality.startswith("Poor"):
            action = "OCR/searchable PDF required for reliable checking"
        rows.append({
            "File": profile.filename,
            "Format": clean(getattr(profile, "document_format", "") or "—").upper(),
            "Detected document": profile.doc_label,
            "Assigned vessel group": assigned_display,
            "Vessel / IMO": vessel,
            "Group confidence": profile.grouping_confidence or "Unassigned",
            "Why grouped": profile.grouping_evidence,
            "Classification confidence": f"{profile.confidence} ({profile.confidence_score}%)",
            "Why classified": profile.classification_evidence,
            "Pages": profile.pages,
            "Text quality": profile.text_quality,
            "Machine-readable": getattr(profile, "machine_readable", "Yes"),
            "Action": action,
        })
    return pd.DataFrame(rows, columns=columns)


def extract_numbered_section(text: str, qid: str, max_chars: int = 16000) -> str:
    """Extract a numbered section without stopping on its own child questions."""
    if not text or not qid:
        return ""
    token = r"\s*\.\s*".join(re.escape(part) for part in qid.split("."))
    match = re.search(r"(?<!\d)" + token + r"(?!\s*\.\s*\d|\d)", text, re.I)
    if not match:
        return ""
    snippet = text[match.start():min(len(text), match.start() + max_chars)]
    target_parts = [int(part) for part in qid.split(".") if part.isdigit()]
    candidate_re = re.compile(r"(?m)^\s*(\d{1,2}(?:\s*\.\s*\d{1,4}){1,3})\b")
    for candidate in candidate_re.finditer(snippet, max(100, len(qid) + 40)):
        parts = [int(p.strip()) for p in re.split(r"\s*\.\s*", candidate.group(1))]
        if len(parts) <= len(target_parts) and parts != target_parts[:len(parts)]:
            snippet = snippet[:candidate.start()]
            break
        if len(parts) == len(target_parts) and parts != target_parts:
            snippet = snippet[:candidate.start()]
            break
    return clean(snippet)


MATERIAL_RE = re.compile(
    r"\b(HMPE|UHMWPE|Dyneema|polyester|polyamide|nylon|polypropylene|mixed\s+polyolefin|"
    r"wire(?:\s+rope)?|steel\s+wire|aramid|Technora|Supermax|Tipto|Atlas)\b",
    re.I,
)


def _nearest_date_context(section: str, keyword_re: str) -> List[Tuple[date, str]]:
    rows: List[Tuple[date, str]] = []
    for match in re.finditer(keyword_re, section or "", re.I):
        window = section[max(0, match.start() - 100):min(len(section), match.end() + 500)]
        for parsed in dates_in_text(window):
            rows.append((parsed, clean(window)[:420]))
    return rows


def _labelled_brake_date_context(section: str) -> List[Tuple[date, str]]:
    """Return only dates tied to an explicit brake/BRC test label.

    Mooring sections often show both a last-test date and a future next-test
    date. Selecting the newest visible date would therefore be unsafe. Prefer
    the first date after a ``last``/``BRC`` label; accept an unlabelled date
    only when it is the sole date in the local context.
    """
    rows: List[Tuple[date, str]] = []
    label_re = re.compile(
        r"(?:date\s+of\s+)?last\s+(?:winch\s+)?(?:brake\s+)?(?:holding\s+capacity\s+)?test|"
        r"(?:BRC|brake\s+holding\s+capacity)\s*(?:test|date)?",
        re.I,
    )
    for match in label_re.finditer(section or ""):
        after = dates_in_text((section or "")[match.end():match.end() + 180])
        if after:
            rows.append((after[0], clean((section or "")[max(0, match.start() - 100):match.end() + 260])[:420]))
    if rows:
        return rows
    # No explicit role: only a single local date is safe to use.
    dates = dates_in_text(section or "")
    if len(dates) == 1:
        return [(dates[0], clean(section or "")[:420])]
    return []


def extract_mooring_summary(texts_by_source: Dict[str, str], ref_date: date) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Return a cautious mooring facts summary and any row-like line/tail details.

    Table layouts vary heavily between HVPQ PDF generators. Values are therefore
    surfaced with confidence and no individual line is called overdue merely from
    a date unless its type/date relationship is visible in the same local context.
    """
    summary: List[Dict[str, Any]] = []
    inventory: List[Dict[str, Any]] = []
    seen_inventory = set()
    for source, text in texts_by_source.items():
        if not text:
            continue
        design_section = extract_numbered_section(text, "10.1.3", 7000)
        brake_section = extract_numbered_section(text, "10.1.4", 9000)
        line_section = extract_numbered_section(text, "10.1.7", 22000)
        combined = " ".join(x for x in [design_section, brake_section, line_section] if x)
        if not combined:
            # Q88 and plans may not use HVPQ numbering.
            matches = [m.start() for m in re.finditer(r"mooring\s+(?:rope|line)|brake\s+(?:test|holding)|SDMBL|LDBF|TDBF", text, re.I)]
            if matches:
                start = max(0, matches[0] - 500)
                combined = clean(text[start:start + 18000])
        if not combined:
            continue

        # Keep the search inside the brake-test section where possible. A long
        # combined section can also contain line installation dates, which must
        # never be presented as the brake-test date.
        brake_dates = _labelled_brake_date_context(brake_section or combined)
        if brake_dates:
            brake_date = max(d for d, _ in brake_dates)
            age_days = (ref_date - brake_date).days
            status = "Within 12-month screening horizon — verify stated interval" if age_days <= 366 else "Older than 12-month screening horizon — verify applicable interval"
            summary.append({
                "Source": source,
                "Item": "Latest visible mooring brake test",
                "Value": brake_date.isoformat(),
                "Status": status,
                "Confidence": "Medium",
                "Evidence": max(brake_dates, key=lambda item: item[0])[1],
            })
        else:
            summary.append({
                "Source": source,
                "Item": "Mooring brake test",
                "Value": "Not reliably extracted",
                "Status": "Needs verification",
                "Confidence": "Low",
                "Evidence": combined[:420],
            })

        for label, pattern in (
            ("SDMBL", r"\bSDMBL\b[^0-9]{0,45}([0-9]+(?:\.[0-9]+)?)\s*(?:t|tonnes?|mt)"),
            ("LDBF", r"\bLDBF\b[^0-9]{0,45}([0-9]+(?:\.[0-9]+)?)\s*(?:t|tonnes?|mt)"),
            ("TDBF", r"\bTDBF\b[^0-9]{0,45}([0-9]+(?:\.[0-9]+)?)\s*(?:t|tonnes?|mt)"),
        ):
            values = list(dict.fromkeys(m.group(1) for m in re.finditer(pattern, combined, re.I)))
            if values:
                summary.append({
                    "Source": source,
                    "Item": label,
                    "Value": ", ".join(values[:12]) + " tonnes",
                    "Status": "Extracted — compare against design and certificates",
                    "Confidence": "Medium",
                    "Evidence": clean(combined[:600]),
                })

        kind_patterns = {
            "Tail": r"\b(?:mooring\s+)?tails?\b",
            "Wire": r"\b(?:mooring\s+)?wire(?:\s+rope)?s?\b",
            "Rope / line": r"\b(?:mooring\s+)?(?:ropes?|lines?)\b",
            "Shackle": r"\bshackles?\b",
        }
        # PDF extraction often flattens the HVPQ table heading into the same
        # line as its first row. Remove the heading/list before looking for row
        # anchors, otherwise its words (ropes, wires, tails and shackles) look
        # like four pieces of equipment.
        item_section = line_section or combined
        item_section = re.sub(
            r"\bprovide\s+details\s+for\b.{0,120}?\bshackles?\b",
            " ",
            item_section,
            flags=re.I,
        )
        item_section = re.sub(
            r"\bmooring\s+ropes?\s*,?\s*wires?\s*,?\s*tails?\s+(?:and|&)\s+shackles?\b",
            " ",
            item_section,
            flags=re.I,
        )

        occurrences: List[Tuple[int, int, str]] = []
        for kind, pattern in kind_patterns.items():
            for match in re.finditer(pattern, item_section, re.I):
                # "wire rope" is one wire record, not a second rope record.
                if kind == "Rope / line" and re.search(r"wire\s*$", item_section[max(0, match.start() - 12):match.start()], re.I):
                    continue
                # A residual prose heading is not inventory evidence.
                if re.search(r"provide\s+details", item_section[max(0, match.start() - 140):match.start()], re.I):
                    continue
                occurrences.append((match.start(), match.end(), kind))
        occurrences.sort(key=lambda item: (item[0], -(item[1] - item[0])))

        # Bound every candidate at the next equipment anchor. This prevents a
        # sparse/blank row from borrowing attributes from a following real row.
        for index, (start, end, kind) in enumerate(occurrences[:60]):
            next_start = occurrences[index + 1][0] if index + 1 < len(occurrences) else min(len(item_section), end + 620)
            window = clean(item_section[start:next_start])
            material = MATERIAL_RE.search(window)
            diameter = re.search(r"\b(\d{2,3}(?:\.\d+)?)\s*mm\b", window, re.I)
            length = re.search(r"\b(\d{2,4}(?:\.\d+)?)\s*(?:m|metres?|meters?)\b", window, re.I)
            strength = re.search(r"\b(?:MBL|SDMBL|LDBF|TDBF|SWL)\b[^0-9]{0,35}([0-9]+(?:\.[0-9]+)?)\s*(?:t|tonnes?|mt)", window, re.I)
            visible_dates = dates_in_text(window)
            location = re.search(r"\b(FWD|FORWARD|AFT|POOP|FORECASTLE|PORT|STBD|STARBOARD|BREAST|SPRING)\b[^,;|]{0,45}", window, re.I)
            # Require at least two attributes so a table heading alone is not presented as a line record.
            attributes = sum(bool(x) for x in [material, diameter, length, strength, visible_dates, location])
            if attributes < 2:
                continue
            item_date = max(visible_dates).isoformat() if visible_dates else ""
            date_basis = ""
            if re.search(r"renew(?:al|ed)|replaced|end[- ]?for[- ]?end|service", window, re.I):
                date_basis = "Renewal/service wording visible"
            elif re.search(r"installed|installation", window, re.I):
                date_basis = "Installation date visible"
            elif item_date:
                date_basis = "Date visible; basis unclear"
            key = (source, kind, material.group(1).lower() if material else "", diameter.group(1) if diameter else "", length.group(1) if length else "", item_date, normalise(location.group(0)) if location else "")
            if key in seen_inventory:
                continue
            seen_inventory.add(key)
            inventory.append({
                "Source": source,
                "Type": kind,
                "Location / identity": clean(location.group(0)) if location else "Not reliably separated",
                "Material": clean(material.group(1)) if material else "",
                "Diameter": f"{diameter.group(1)} mm" if diameter else "",
                "Length": f"{length.group(1)} m" if length else "",
                "Strength value": f"{strength.group(1)} tonnes" if strength else "",
                "Visible date": item_date,
                "Last renewal/service date": item_date,
                "Date basis": date_basis,
                "Confidence": "Medium" if attributes >= 3 else "Low",
                "Evidence excerpt": window[:500],
            })

        counts = pd.DataFrame(inventory)
        if not counts.empty:
            local = counts[counts["Source"] == source]
            if not local.empty:
                for kind, count in local["Type"].value_counts().items():
                    summary.append({
                        "Source": source,
                        "Item": f"{kind} records recognised",
                        "Value": str(int(count)),
                        "Status": "Partial inventory — confirm against every HVPQ/LMP row",
                        "Confidence": "Medium",
                        "Evidence": "Structured from local row context; table formatting may merge cells.",
                    })

    summary_cols = ["Source", "Item", "Value", "Status", "Confidence", "Evidence"]
    inventory_cols = ["Source", "Type", "Location / identity", "Material", "Diameter", "Length", "Strength value", "Visible date", "Last renewal/service date", "Date basis", "Confidence", "Evidence excerpt"]
    return pd.DataFrame(summary, columns=summary_cols).drop_duplicates(), pd.DataFrame(inventory, columns=inventory_cols).drop_duplicates()


def _field_score(field: Any) -> int:
    score = int(getattr(field, "confidence_score", 0) or 0)
    if not score:
        label = clean(getattr(field, "confidence", "")).lower()
        score = {"table-aware": 96, "deterministic": 92, "section-snippet": 78, "best-effort": 68, "local-llm": 62, "manual-needed": 30}.get(label, 70)
    return score + SOURCE_PRIORITY.get(clean(getattr(field, "source", "")).upper(), 0)


def best_field(fields: Sequence[Any], field_id: str, sources: Optional[Sequence[str]] = None) -> Optional[Any]:
    candidates = [f for f in fields if getattr(f, "field_id", "") == field_id and clean(getattr(f, "value", ""))]
    if sources is not None:
        allowed = {s.upper() for s in sources}
        candidates = [f for f in candidates if clean(getattr(f, "source", "")).upper() in allowed]
    if not candidates:
        return None
    return max(candidates, key=_field_score)


def certificate_watch_df(fields: Sequence[Any], ref_date: date, horizon_days: int = 180) -> pd.DataFrame:
    keys = set()
    for field in fields:
        match = re.fullmatch(r"cert\.([^.]+)\.expiry", clean(getattr(field, "field_id", "")))
        if match:
            keys.add(match.group(1))
    rows: List[Dict[str, Any]] = []
    for key in sorted(keys):
        candidates = [f for f in fields if getattr(f, "field_id", "") == f"cert.{key}.expiry" and parse_date(getattr(f, "value", ""))]
        if not candidates:
            continue
        preferred = max(candidates, key=_field_score)
        expiry = parse_date(getattr(preferred, "value", ""))
        if not expiry:
            continue
        days = (expiry - ref_date).days
        if days < 0:
            status, urgency = "Expired", "Immediate hold / verify"
        elif days <= 30:
            status, urgency = "Expires within 30 days", "Immediate attention"
        elif days <= 90:
            status, urgency = "Short-term — expires within 90 days", "Before fixture / voyage"
        elif days <= horizon_days:
            status, urgency = f"Watch — expires within {horizon_days} days", "Plan renewal"
        else:
            status, urgency = "Valid beyond watch window", "No immediate action"
        by_source = {clean(getattr(f, "source", "")): clean(getattr(f, "value", "")) for f in candidates}
        preferred_file = clean(getattr(preferred, "document_name", ""))
        preferred_page = int(getattr(preferred, "page", 0) or 0)
        evidence_location = preferred_file + (f" p.{preferred_page}" if preferred_page else "")
        rows.append({
            "Certificate": CERTIFICATE_NAMES.get(key, key.replace("_", " ").title()),
            "Status": status,
            "Days remaining": days,
            "Expiry used": expiry.isoformat(),
            "Preferred source": clean(getattr(preferred, "source", "")),
            "Evidence file / page": evidence_location,
            "HVPQ": by_source.get("HVPQ", ""),
            "Class Status": by_source.get("CLASS", ""),
            "Certificate file": by_source.get("CERTIFICATE", ""),
            "Action": urgency,
            "Confidence": _confidence_label(min(99, _field_score(preferred))),
        })
    columns = ["Certificate", "Status", "Days remaining", "Expiry used", "Preferred source", "Evidence file / page", "HVPQ", "Class Status", "Certificate file", "Action", "Confidence"]
    df = pd.DataFrame(rows, columns=columns)
    if df.empty:
        return df
    order = {"Expired": 0, "Expires within 30 days": 1, "Short-term — expires within 90 days": 2}
    df["_rank"] = df["Status"].map(lambda x: order.get(x, 3 if str(x).startswith("Watch") else 4))
    return df.sort_values(["_rank", "Days remaining", "Certificate"]).drop(columns="_rank")


def enhance_register(df: pd.DataFrame) -> pd.DataFrame:
    """Add a decision layer without relabelling extraction gaps as errors."""
    if df is None or df.empty:
        return df
    out = df.copy()
    conclusions: List[str] = []
    confidence: List[str] = []
    user_meaning: List[str] = []
    for _, row in out.iterrows():
        joined = " ".join(clean(v).lower() for v in row.values)
        priority = clean(row.get("Priority", "")).upper()
        status = clean(row.get("Status", "")).lower()
        has_gap = any(token in joined for token in ["not reliably extracted", "not extracted", "could not reliably", "section not", "blank/not extracted", "manual confirmation"])
        uncertain = any(token in joined for token in ["appears", "may ", "potential", "latest visible", "verify whether", "could be"])
        if priority == "OK" or "in order" in status:
            conclusion, conf, meaning = "In order from available evidence", "High" if not uncertain else "Medium", "No correction indicated; retain evidence"
        elif "manual" in priority.lower() or "manual" in status or has_gap:
            conclusion, conf, meaning = "Not verified", "Low", "Check the source document; this is not a proven defect"
        elif "mismatch" in status or "not satisfactory" in status or priority in {"CRITICAL", "HIGH", "MEDIUM"}:
            if uncertain:
                conclusion, conf, meaning = "Potential concern", "Medium", "Verify promptly before treating as a deficiency"
            else:
                conclusion, conf, meaning = "Document-supported discrepancy", "High", "Correct or formally resolve with source evidence"
        else:
            conclusion, conf, meaning = "Review", "Medium", "Review evidence and decide"
        conclusions.append(conclusion)
        confidence.append(conf)
        user_meaning.append(meaning)
    out.insert(0, "Conclusion", conclusions)
    out.insert(1, "Evidence confidence", confidence)
    out.insert(2, "What this means", user_meaning)
    rank = {
        "Document-supported discrepancy": 0,
        "Potential concern": 1,
        "Not verified": 2,
        "Review": 3,
        "In order from available evidence": 4,
    }
    out["_decision_rank"] = out["Conclusion"].map(rank).fillna(9)
    return out.sort_values(["_decision_rank"]).drop(columns="_decision_rank")


def conclusion_counts(*frames: pd.DataFrame) -> Dict[str, int]:
    counts = {"confirmed": 0, "potential": 0, "unverified": 0, "ok": 0}
    for frame in frames:
        if frame is None or frame.empty or "Conclusion" not in frame.columns:
            continue
        series = frame["Conclusion"].astype(str)
        counts["confirmed"] += int(series.str.contains("Document-supported", case=False, na=False).sum())
        counts["potential"] += int(series.str.contains("Potential", case=False, na=False).sum())
        counts["unverified"] += int(series.str.contains("Not verified", case=False, na=False).sum())
        counts["ok"] += int(series.str.contains("In order", case=False, na=False).sum())
    return counts


def extract_sts_particulars(text: str) -> Dict[str, str]:
    patterns = {
        "LOA": r"\b(?:LOA|Length\s+Overall)\b[^0-9]{0,45}([0-9]{2,3}(?:\.[0-9]+)?)\s*(?:m|metres?|meters?)",
        "Beam": r"\b(?:Beam|Breadth(?:\s+Moulded)?)\b[^0-9]{0,45}([0-9]{1,3}(?:\.[0-9]+)?)\s*(?:m|metres?|meters?)",
        "Summer draft": r"\b(?:Summer\s+(?:draft|draught)|Draft\s+Summer)\b[^0-9]{0,45}([0-9]{1,2}(?:\.[0-9]+)?)\s*(?:m|metres?|meters?)",
        "DWT": r"\b(?:DWT|Deadweight)\b[^0-9]{0,45}([0-9][0-9, ]{3,12}(?:\.[0-9]+)?)\s*(?:MT|tonnes?|t)?",
        "Manifold height": r"\b(?:manifold\s+height|height\s+of\s+(?:cargo\s+)?manifold)\b[^0-9]{0,70}([0-9]{1,2}(?:\.[0-9]+)?)\s*(?:m|metres?|meters?)",
        "Maximum transfer rate": r"\b(?:maximum|max\.?)[^\n]{0,45}(?:loading|discharg(?:e|ing)|transfer)\s+rate\b[^0-9]{0,45}([0-9][0-9, ]{2,10})\s*(?:m3|m³)(?:/h|/hr|\s+per\s+hour)",
        "Maximum manifold pressure": r"\b(?:maximum|max\.?)\s+(?:allowable\s+)?(?:manifold\s+)?pressure\b[^0-9]{0,45}([0-9]{1,3}(?:\.[0-9]+)?)\s*(bar|kg/cm2|kpa)",
        "SDMBL": r"\bSDMBL\b[^0-9]{0,45}([0-9]+(?:\.[0-9]+)?)\s*(?:t|tonnes?|mt)",
    }
    result: Dict[str, str] = {}
    for label, pattern in patterns.items():
        match = re.search(pattern, text or "", re.I | re.S)
        if not match:
            continue
        unit = ""
        if label in {"LOA", "Beam", "Summer draft", "Manifold height"}:
            unit = " m"
        elif label == "DWT":
            unit = " tonnes"
        elif label == "Maximum transfer rate":
            unit = " m³/h"
        elif label == "SDMBL":
            unit = " tonnes"
        elif label == "Maximum manifold pressure" and len(match.groups()) > 1:
            unit = " " + match.group(2)
        result[label] = clean(match.group(1)).replace(" ", "") + unit

    sizes = []
    for match in re.finditer(r"(?:cargo|vapou?r|bunker)?\s*manifold[^\n]{0,180}?\b(\d{1,2})\s*(?:inches?|inch|in\.?|\")", text or "", re.I):
        sizes.append(match.group(1) + " in")
    if sizes:
        result["Visible manifold sizes"] = ", ".join(dict.fromkeys(sizes))
    return result


def sts_compatibility_facts(primary: Dict[str, str], counterpart: Dict[str, str], primary_label: str, counterpart_label: str) -> pd.DataFrame:
    fact_names = ["LOA", "Beam", "Summer draft", "DWT", "Manifold height", "Visible manifold sizes", "Maximum transfer rate", "Maximum manifold pressure", "SDMBL"]
    rows = []
    for fact in fact_names:
        a, b = primary.get(fact, ""), counterpart.get(fact, "")
        result = "Values extracted — apply operation-specific compatibility limits"
        if not a or not b:
            result = "Not verified — value missing for one or both vessels"
        elif fact == "Visible manifold sizes":
            aset = {x.strip() for x in a.split(",")}
            bset = {x.strip() for x in b.split(",")}
            result = "Common nominal size visible: " + ", ".join(sorted(aset & bset)) + " — verify standard/reducers/hoses" if aset & bset else "No common visible size — verify reducers/hoses"
        rows.append({"Compatibility fact": fact, primary_label: a or "Not extracted", counterpart_label: b or "Not extracted", "Screening result": result})
    return pd.DataFrame(rows)


def sts_document_check_rows(profiles: Sequence[DocumentProfile], vessel_imo: str, vessel_label: str) -> List[Dict[str, str]]:
    vessel_profiles = []
    for profile in profiles:
        related = {value.strip() for value in clean(getattr(profile, "related_imos", "")).split(",") if value.strip()}
        shared_assessment = profile.doc_type == "STS_ASSESSMENT" and vessel_imo in related
        if profile.assigned_group == vessel_imo or profile.imo == vessel_imo or shared_assessment:
            vessel_profiles.append(profile)
    types = {p.doc_type for p in vessel_profiles}
    requirements = [
        ("Vessel particulars", {"HVPQ", "Q88"}, "Required to establish principal particulars and compatibility inputs"),
        ("Class status", {"CLASS"}, "Required to identify Conditions/Memoranda and survey status"),
        ("P&I / insurance", {"PNI"}, "Required to confirm current cover and relevant liabilities"),
        ("Certificate evidence", {"CERTIFICATE"}, "Original/current certificate evidence should support extracted dates"),
        ("STS plan / procedure", {"STS_PLAN"}, "Required onboard operational control document"),
        ("Operation-specific compatibility / JPO / risk assessment", {"STS_ASSESSMENT"}, "Required before final operation clearance"),
        ("Mooring management evidence", {"MOORING_PLAN", "HVPQ"}, "Needed to verify line/tail/brake suitability and records"),
    ]
    rows: List[Dict[str, str]] = []
    for item, alternatives, why in requirements:
        present = sorted(types & alternatives)
        if present:
            # Detecting one certificate PDF does not prove that the complete
            # statutory/trading set is present and current.
            status = "REVIEW" if item == "Certificate evidence" else "PASS"
            evidence = ", ".join(DOC_LABELS.get(x, x) for x in present)
            action = "Confirm the full required certificate set and voyage coverage" if item == "Certificate evidence" else "No document gap from uploaded set"
            confidence = "High"
        else:
            status = "NOT VERIFIED"
            evidence = "No matching document identified"
            action = "Obtain/review before final STS clearance"
            confidence = "High"
        rows.append({
            "Status": status,
            "Vessel": vessel_label,
            "Area": "Document readiness",
            "Check": item,
            "Evidence": evidence,
            "Why it matters": why,
            "Required action": action,
            "Confidence": confidence,
        })
    return rows


def build_sts_screen(
    profiles: Sequence[DocumentProfile],
    fields_by_vessel: Dict[str, Sequence[Any]],
    texts_by_vessel: Dict[str, Sequence[str]],
    primary_imo: str,
    counterpart_imo: str,
    ref_date: date,
    watch_days: int = 180,
) -> Tuple[str, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    def label_for(imo: str) -> str:
        names = [p.vessel_name for p in profiles if (p.assigned_group == imo or p.imo == imo) and p.vessel_name]
        return f"{names[0]} / IMO {imo}" if names else f"IMO {imo}"

    primary_label, counterpart_label = label_for(primary_imo), label_for(counterpart_imo)
    rows = sts_document_check_rows(profiles, primary_imo, primary_label) + sts_document_check_rows(profiles, counterpart_imo, counterpart_label)
    cert_frames: List[pd.DataFrame] = []
    for imo, vessel_label in ((primary_imo, primary_label), (counterpart_imo, counterpart_label)):
        fields = list(fields_by_vessel.get(imo, []))
        watch = certificate_watch_df(fields, ref_date, horizon_days=watch_days)
        if not watch.empty:
            watch = watch.copy()
            watch.insert(0, "Vessel", vessel_label)
            cert_frames.append(watch)
            for _, item in watch.iterrows():
                status_text = clean(item["Status"])
                if status_text == "Expired":
                    status, action = "HOLD", "Verify/renew certificate before clearance"
                elif "30 days" in status_text or "90 days" in status_text:
                    status, action = "REVIEW", "Confirm validity covers the intended operation and voyage"
                else:
                    continue
                rows.append({
                    "Status": status,
                    "Vessel": vessel_label,
                    "Area": "Certificates",
                    "Check": clean(item["Certificate"]),
                    "Evidence": f"Expiry {item['Expiry used']} ({item['Days remaining']} days remaining); source {item.get('Evidence file / page', '') or item['Preferred source']}",
                    "Why it matters": "Expired or short-validity certification can prevent fixture or operational acceptance.",
                    "Required action": action,
                    "Confidence": clean(item["Confidence"]),
                })

        conditions = best_field(fields, "classification.conditions_of_class", ["CLASS", "HVPQ", "Q88"])
        memoranda = best_field(fields, "classification.memo_of_class", ["CLASS", "HVPQ", "Q88"])
        dispensation = best_field(fields, "classification.flag_dispensation", ["CLASS", "HVPQ", "Q88"])
        for title, field, severity in (("Conditions of Class", conditions, "HOLD"), ("Memoranda / recommendations", memoranda, "REVIEW"), ("Flag/Class dispensation", dispensation, "REVIEW")):
            value = clean(getattr(field, "value", "")) if field else ""
            nvalue = normalise(value)
            if value and nvalue not in {"no", "none", "nil", "na", "not applicable", "0", "zero"}:
                rows.append({
                    "Status": severity,
                    "Vessel": vessel_label,
                    "Area": "Class / statutory",
                    "Check": title,
                    "Evidence": value,
                    "Why it matters": "This item may affect acceptability, operational limitations or approval conditions.",
                    "Required action": "Review the full entry, due date and operational restriction before clearance.",
                    "Confidence": _confidence_label(min(99, _field_score(field))) if field else "Low",
                })
            elif not value:
                rows.append({
                    "Status": "NOT VERIFIED",
                    "Vessel": vessel_label,
                    "Area": "Class / statutory",
                    "Check": title,
                    "Evidence": "Not reliably extracted",
                    "Why it matters": "Absence cannot be assumed to mean Nil.",
                    "Required action": "Confirm against current Class Status / dispensation evidence.",
                    "Confidence": "Low",
                })

        detained = best_field(fields, "psc.detained_36m", ["PIQ", "HVPQ", "Q88"])
        if detained and normalise(getattr(detained, "value", "")) in {"yes", "y", "true", "1"}:
            rows.append({
                "Status": "REVIEW", "Vessel": vessel_label, "Area": "PSC / history", "Check": "PSC detention declared",
                "Evidence": clean(getattr(detained, "value", "")), "Why it matters": "Recent detention may trigger charterer or terminal review.",
                "Required action": "Review detention report, closure evidence and acceptance criteria.", "Confidence": "High",
            })

    primary_text = "\n".join(texts_by_vessel.get(primary_imo, []))
    counterpart_text = "\n".join(texts_by_vessel.get(counterpart_imo, []))
    facts = sts_compatibility_facts(
        extract_sts_particulars(primary_text), extract_sts_particulars(counterpart_text), primary_label, counterpart_label
    )
    if not facts.empty:
        for _, fact in facts.iterrows():
            result = clean(fact["Screening result"])
            if result.startswith("No common"):
                status = "REVIEW"
            elif result.startswith("Not verified"):
                status = "NOT VERIFIED"
            else:
                # Extracted dimensions/rates are compatibility inputs, not an
                # automatic compatibility decision without operation-specific
                # limits, fender/geometry review and responsible approval.
                status = "REVIEW"
            rows.append({
                "Status": status,
                "Vessel": "Both vessels",
                "Area": "Compatibility particulars",
                "Check": clean(fact["Compatibility fact"]),
                "Evidence": f"{primary_label}: {fact[primary_label]} | {counterpart_label}: {fact[counterpart_label]}",
                "Why it matters": "The values feed the formal compatibility study and operation-specific risk assessment.",
                "Required action": "Use verified figures in the compatibility study/JPO." if status != "PASS" else "Carry verified values into the final compatibility study.",
                "Confidence": "Medium" if status == "PASS" else "Low",
            })

    screen_columns = ["Status", "Vessel", "Area", "Check", "Evidence", "Why it matters", "Required action", "Confidence"]
    screen = pd.DataFrame(rows, columns=screen_columns).drop_duplicates()
    status_rank = {"HOLD": 0, "REVIEW": 1, "NOT VERIFIED": 2, "PASS": 3}
    if not screen.empty:
        screen["_rank"] = screen["Status"].map(status_rank).fillna(9)
        screen = screen.sort_values(["_rank", "Vessel", "Area", "Check"]).drop(columns="_rank")
    decision = "HOLD" if (screen["Status"] == "HOLD").any() else ("CONDITIONAL / INCOMPLETE" if screen["Status"].isin(["REVIEW", "NOT VERIFIED"]).any() else "PRELIMINARY CLEAR")
    cert_watch = pd.concat(cert_frames, ignore_index=True) if cert_frames else pd.DataFrame()
    return decision, screen, facts, cert_watch
