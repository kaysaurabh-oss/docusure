from __future__ import annotations

import io, re, zipfile
from dataclasses import dataclass, asdict
from datetime import date
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import streamlit as st
from dateutil import parser as dateparser
from dateutil.relativedelta import relativedelta

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

APP_VERSION = "v8-extraction-first"

# -----------------------------
# Models
# -----------------------------
@dataclass
class Field:
    value: str = ""
    source: str = ""
    evidence: str = ""
    confidence: str = "HIGH"

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
    reason: str = ""
    required_action: str = ""
    evidence: str = ""

# Month-name + ISO dates. Avoid numeric-only dates to stop qids becoming dates.
DATE_RE = re.compile(r"(?:\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|[A-Za-z]{3,9}\s+\d{1,2},\s*\d{4}|\d{4}-\d{2}-\d{2}|\d{1,2}[./-][A-Za-z]{3,9}[./-]\d{4})", re.I)

CERT_ALIASES = {
    "safety_equipment": ["Safety Equipment Certificate", "Safety Equipment Certificate (SEC)", "Cargo Ship Safety Equipment Certificate"],
    "safety_radio": ["Safety Radio Certificate", "Safety Radio Certificate (SRC)", "Cargo Ship Safety Radio Certificate"],
    "safety_construction": ["Safety Construction Certificate", "Safety Construction Certificate (SCC)", "Cargo Ship Safety Construction Certificate"],
    "loadline": ["Loadline Certificate", "Load Line Certificate", "International Loadline Certificate", "International Loadline Certificate (ILC)"],
    "iopp": ["International Oil Pollution Prevention Certificate", "International Oil Pollution Prevention Certificate (IOPPC)", "Oil Pollution Prevention Certificate(Form B)", "OPP (MARPOL Annex I)"],
    "ibwmc": ["International Ballast Water Management Certificate", "International Ballast Water Management Certificate (IBWMC)", "Ballast Water Management Certificate"],
    "smc": ["Safety Management Certificate", "ISM Safety Management Certificate", "Safety Management Certificate (SMC)"],
    "doc": ["Document of Compliance", "Document of Compliance (DOC)"],
    "issc": ["International Ship Security Certificate", "International Ship Security Certificate (ISSC)"],
    "uscg_coc": ["USCG Certificate of Compliance", "USCG Certificate of Compliance(USCGCOC)", "USCG Requirement for Pollution Prevention"],
    "cofr": ["U.S. Certificate of Financial Responsibility", "U.S. Certificate of Financial Responsibility - Expiry Date", "Certificate of Financial Responsibility", "COFR"],
    "vgp": ["Vessel General Permit", "Vessel General Permit Issue date"],
    "cof_chem": ["Certificate of Fitness (Expiry Dates) - Chemicals", "Certificate of Fitness (COF) (Chemical)", "Bulk Chemical Code Certificate", "Certificate of Fitness"],
    "class_certificate": ["Class Certificate", "Certificate of Class (COC)", "Classification Certificate"],
    "clc_oil": ["Civil Liability Convention Certificate (1992)", "Civil Liability Convention (CLC) 1992 Certificate", "Civil Liability Certificates"],
    "clc_bunker": ["Civil Liability Convention 2001", "Civil Liability for Bunker Oil Pollution Damage Convention", "Bunker"],
    "wreck_removal": ["Wreck removal Convention Certificate", "Liability for the Removal of Wrecks Certificate", "WRC"],
}

CERT_ORDER = list(CERT_ALIASES.keys())

# -----------------------------
# Helpers
# -----------------------------
def norm_space(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()

def clean(s: Any) -> str:
    return norm_space(s).strip(" :;,.—-")

def pdf_text(uploaded) -> str:
    if uploaded is None or fitz is None:
        return ""
    data = uploaded.getvalue() if hasattr(uploaded, "getvalue") else uploaded.read()
    doc = fitz.open(stream=data, filetype="pdf")
    pages = []
    for i, p in enumerate(doc, start=1):
        pages.append(f"\n--- PAGE {i} ---\n" + p.get_text("text"))
    return "\n".join(pages)

def dates_in(s: str) -> List[date]:
    out=[]
    for m in DATE_RE.finditer(s or ""):
        raw=m.group(0)
        try:
            if re.match(r"^\d{4}-\d{2}-\d{2}$", raw):
                out.append(date.fromisoformat(raw))
            else:
                out.append(dateparser.parse(raw, dayfirst=True).date())
        except Exception:
            pass
    return out

def dstr(d: date|str|None) -> str:
    if isinstance(d, date): return d.isoformat()
    return str(d or "")

def setf(d: Dict[str, Field], key: str, value: Any, source: str, evidence: str, confidence="HIGH"):
    val = clean(value)
    if not val:
        return
    if key not in d or not d[key].value or (d[key].confidence != "HIGH" and confidence == "HIGH"):
        d[key] = Field(val, source, norm_space(evidence)[:1500], confidence)

def block_between(text: str, start_pat: str, end_pat: Optional[str]=None, flags=re.I|re.S, max_chars=4000) -> str:
    m = re.search(start_pat, text or "", flags)
    if not m: return ""
    start = m.start()
    sub = text[start:start+max_chars]
    if end_pat:
        e = re.search(end_pat, sub[len(m.group(0)):], flags)
        if e:
            sub = sub[:len(m.group(0))+e.start()]
    return sub

def qblock(text: str, qid: str, max_lines=70) -> str:
    lines=[clean(x) for x in (text or "").splitlines()]
    start=None
    for i,ln in enumerate(lines):
        if ln == qid or ln.startswith(qid+" ") or ln.startswith(qid+"."):
            start=i; break
    if start is None:
        pat=re.compile(rf"\b{re.escape(qid)}\b")
        for i,ln in enumerate(lines):
            if pat.search(ln): start=i; break
    if start is None: return ""
    out=[]
    next_re=re.compile(r"^\d+(?:\.\d+){1,4}[a-z]?\.?$")
    for j in range(start,min(len(lines),start+max_lines)):
        if j>start and next_re.match(lines[j]):
            break
        if lines[j]: out.append(lines[j])
    return " ".join(out)

def yn_from_block(b: str) -> str:
    vals=re.findall(r"\b(Yes|No|Not applicable|N/A|NA|Nil)\b", b or "", re.I)
    if not vals: return ""
    v=vals[-1].lower()
    if v in {"n/a","na","not applicable","nil"}: return "No"
    return "Yes" if v=="yes" else "No"

def line_answer_after_label(text: str, label_regex: str, stop_regex: Optional[str]=None) -> str:
    m=re.search(label_regex + r"\s*[:?]?\s*(.+)", text or "", re.I)
    if m:
        val=m.group(1)
        if stop_regex:
            val=re.split(stop_regex,val,flags=re.I)[0]
        return clean(val)
    return ""

def norm_compare(s: str) -> str:
    x=(s or "").lower()
    x=re.sub(r"[^a-z0-9]+"," ",x)
    x=norm_space(x)
    reps={
        "not applicable":"no", "n a":"no", "nil":"no", "none":"no",
        "korean register of shipping":"korean register", "kr":"korean register",
        "northstandard limitied":"northstandard limited",
        "p i":"pi",
    }
    for k,v in reps.items(): x=x.replace(k,v)
    return norm_space(x)

def same_date(a: str, b: str) -> Optional[bool]:
    da,db=dates_in(a),dates_in(b)
    if da and db: return da[0]==db[0]
    # ISO fields
    try:
        if re.match(r"\d{4}-\d{2}-\d{2}$", a) and re.match(r"\d{4}-\d{2}-\d{2}$", b):
            return a==b
    except Exception: pass
    return None

def equivalent(a: str,b: str, field="") -> bool:
    if not a or not b: return True
    sd=same_date(a,b)
    if sd is not None: return sd
    na,nb=norm_compare(a),norm_compare(b)
    if na==nb: return True
    if field=="vessel_name" and (na in nb or nb in na): return True
    # Ship type is deliberately loose: different forms use different wording.
    if field=="vessel_type":
        family_words=["oil", "chemical", "product", "crude", "tanker", "carrier"]
        ca={w for w in family_words if w in na}
        cb={w for w in family_words if w in nb}
        if ca & cb and not ({"gas","bulk"} & (ca|cb)):
            return True
    if field=="class_notation":
        tokens=["iws","igs","cow","bwm","ihm","lg","vec","esp"]
        a_tokens={t for t in tokens if t in na}
        b_tokens={t for t in tokens if t in nb}
        return len(a_tokens & b_tokens) >= min(3, len(a_tokens), len(b_tokens))
    if len(na)>8 and len(nb)>8 and (na in nb or nb in na): return True
    return False

def cert_segment(text: str, aliases: List[str], all_aliases: List[str], max_chars=700) -> str:
    # Normalize whitespace first so wrapped labels like "U.S. Certificate of\nFinancial Responsibility" are found.
    norm = norm_space(text)
    lower = norm.lower()
    positions=[]
    for a in aliases:
        i=lower.find(norm_space(a).lower())
        if i!=-1: positions.append(i)
    if not positions: return ""
    start=min(positions)
    stop=start+max_chars
    alias_lowers=[norm_space(x).lower() for x in aliases]
    for a in all_aliases:
        aa=norm_space(a).lower()
        if aa in alias_lowers:
            continue
        j=lower.find(aa, start+5)
        if j!=-1 and j<stop: stop=j
    return norm[start:stop]

def add_cert_field(fields: Dict[str, Field], key: str, prefix: str, issue: str, expiry: str, source: str, seg: str, annual="", intermediate=""):
    if issue: setf(fields, f"cert_{key}_issue", issue, source, seg)
    if expiry: setf(fields, f"cert_{key}_expiry", expiry, source, seg)
    if annual: setf(fields, f"cert_{key}_annual", annual, source, seg)
    if intermediate: setf(fields, f"cert_{key}_intermediate", intermediate, source, seg)

# -----------------------------
# HVPQ extraction
# -----------------------------
def extract_hvpq(text: str) -> Dict[str, Field]:
    f={}
    if not text: return f
    m=re.search(r"Harmonised Vessel Particulars Questionnaire v6\s+([0-9]{1,2}\s+[A-Za-z]+\s+\d{4}).*?IMO/LR Number\s+(\d{7})\s+([A-Z0-9 '\-]+)", text, re.S)
    if m:
        setf(f,"doc_date",dstr(dates_in(m.group(1))[0]),"HVPQ",m.group(0))
        setf(f,"imo_number",m.group(2),"HVPQ",m.group(0))
        setf(f,"vessel_name",m.group(3),"HVPQ",m.group(0))
    # General fields using qblocks
    for qid,key in [("1.1.4","flag"),("1.1.5","port_registry"),("1.1.6","call_sign"),("1.1.10","mmsi"),("1.1.8","vessel_type"),("1.1.9","vessel_type_other"),("1.1.13","pni_club")]:
        b=qblock(text,qid,45)
        if not b: continue
        if key=="flag":
            m=re.search(r"1 Flag\s+(.+?)(?:\s+2 Has|$)", b, re.I)
            if m: setf(f,key,re.sub(r"^\?\s*", "", m.group(1)),"HVPQ",b)
        elif key=="vessel_type":
            m=re.search(r"IOPPC\)?\s*(Oil Tanker|Other|Product Carrier|Chemical Tanker|.+?)(?:\s+1\.1\.9|$)", b, re.I)
            if m: setf(f,key,m.group(1),"HVPQ",b)
        elif key=="vessel_type_other":
            m=re.search(r"specify\s+(.+?)(?:\s+1\.1\.10|$)", b, re.I)
            if m: setf(f,key,m.group(1),"HVPQ",b)
        elif key=="pni_club":
            m=re.search(r"If other, then specify\s+(.+?)(?:\s+3 Amount|$)", b, re.I)
            if m: setf(f,key,m.group(1),"HVPQ",b)
    # Owner/operator
    b=qblock(text,"1.3.1",90)
    m=re.search(r"1 Name\s+(.+?)(?:\s+2 Full|$)",b,re.I)
    if m: setf(f,"registered_owner",m.group(1),"HVPQ",b)
    b=qblock(text,"1.3.2",90)
    m=re.search(r"1 Name\s+(.+?)(?:\s+2 Full|$)",b,re.I)
    if m: setf(f,"technical_operator",m.group(1),"HVPQ",b)
    # Energy
    b=qblock(text,"1.2.3",80)
    m=re.search(r"provide CII rating\s+([A-E])",b,re.I)
    if m: setf(f,"cii_rating",m.group(1),"HVPQ",b)
    m=re.search(r"CII rating verified by Class, 3rd Party or Owner\??\s*(Class|Owner|3rd Party)",b,re.I)
    if m: setf(f,"cii_verified_by",m.group(1),"HVPQ",b)
    # Classification/surveys
    b=qblock(text,"1.5.1",80)
    m=re.search(r"1 Classification Society\s+(.+?)(?:\s+2 Is|$)",b,re.I)
    if m:
        val=re.sub(r"^1\s+Classification Society\s+", "", m.group(1), flags=re.I)
        setf(f,"class_society",val,"HVPQ",b)
    b=qblock(text,"1.5.2",120)
    m=re.search(r"List class notations\s+(.+?)(?:\s+2 Provide|$)",b,re.I)
    if m: setf(f,"class_notation",m.group(1),"HVPQ",b)
    b=qblock(text,"1.5.4",80); ds=dates_in(b)
    if len(ds)>=1: setf(f,"last_drydock",dstr(ds[0]),"HVPQ",b)
    if len(ds)>=3: setf(f,"next_drydock_due",dstr(ds[2]),"HVPQ",b)
    b=qblock(text,"1.5.5",50); ds=dates_in(b)
    if len(ds)>=1: setf(f,"last_iws",dstr(ds[0]),"HVPQ",b)
    if len(ds)>=2: setf(f,"next_iws_due",dstr(ds[1]),"HVPQ",b)
    b=qblock(text,"1.5.6",100); ds=dates_in(b)
    if len(ds)>=1: setf(f,"last_special_survey",dstr(ds[0]),"HVPQ",b)
    if len(ds)>=2: setf(f,"next_special_survey_due",dstr(ds[-1]),"HVPQ",b)
    b=qblock(text,"1.5.11",25); ds=dates_in(b)
    if ds: setf(f,"last_annual_survey",dstr(ds[0]),"HVPQ",b)
    b=qblock(text,"1.5.12",25); ds=dates_in(b)
    if ds: setf(f,"last_intermediate_survey",dstr(ds[0]),"HVPQ",b)
    for qid,key in [("1.5.14","conditions_of_class"),("1.5.16","memoranda_of_class"),("1.5.18","flag_dispensation")]:
        b=qblock(text,qid,35); yn=yn_from_block(b)
        if yn: setf(f,key,yn,"HVPQ",b)
    # Incidents/PSC
    b=qblock(text,"1.9.3",45); yn=yn_from_block(b)
    if yn: setf(f,"incident_pollution_grounding_collision",yn,"HVPQ",b)
    b=qblock(text,"1.9.5",70); yn=yn_from_block(b)
    if yn: setf(f,"incident_other",yn,"HVPQ",b)
    b=qblock(text,"1.9.8",80); ds=dates_in(b)
    if ds: setf(f,"last_psc_date",dstr(ds[0]),"HVPQ",b)
    m=re.search(r"Port of last Port State Control Inspection\s+(.+?)(?:\s+3 Has|$)",b,re.I)
    if m: setf(f,"last_psc_port",m.group(1),"HVPQ",b)
    m=re.search(r"detained.*?\?\s*(Yes|No)",b,re.I)
    if m: setf(f,"psc_detained",m.group(1),"HVPQ",b)
    # Certificates section
    cert_text=block_between(text, r"2\.1\.5\s*Certificate dates|Date Issued\s+Date Expires", r"2\.2\.1\s*Publications|Publications", max_chars=7000)
    if not cert_text:
        cert_text=block_between(text, r"Date Issued\s+Date Expires", r"Publications", max_chars=7000)
    all_aliases=[a for v in CERT_ALIASES.values() for a in v]
    for key, aliases in CERT_ALIASES.items():
        seg=cert_segment(cert_text, aliases, all_aliases, 900)
        if not seg: continue
        ds=dates_in(seg)
        if not ds: continue
        if key=="vgp":
            # HVPQ sometimes carries issue date + next date; keep both.
            issue=dstr(ds[0]); expiry=dstr(ds[-1]) if len(ds)>1 else ""
        else:
            issue=dstr(ds[0])
            expiry=dstr(ds[1]) if len(ds)>1 else ""
        annual=dstr(ds[2]) if len(ds)>2 else ""
        inter=dstr(ds[3]) if len(ds)>3 else ""
        add_cert_field(f,key,"cert",issue,expiry,"HVPQ",seg,annual,inter)
    # Publications section for targeted check.
    pubs=block_between(text, r"2\.2\.1\s*Publications|Publications\s+Edition Number", r"3\s+Crew|3\.1\.1", max_chars=4500)
    if pubs: setf(f,"publications_block","Present","HVPQ",pubs,"MEDIUM")
    return f

# -----------------------------
# PIQ extraction
# -----------------------------
def extract_piq(text: str) -> Dict[str, Field]:
    f={}
    if not text: return f
    m=re.search(r"PIQ Report\s+(?:Vessel Name\s+)?([A-Z0-9 '\-]+)\s+Date\s+([0-9]{1,2}\s+[A-Za-z]+\s+\d{4})",text,re.S)
    if m:
        setf(f,"vessel_name",m.group(1),"PIQ",m.group(0))
        setf(f,"doc_date",dstr(dates_in(m.group(2))[0]),"PIQ",m.group(0))
    # PIQ table extraction is layout-driven: answer often appears before the label in PyMuPDF text.
    b=block_between(text, r"1\.1\.1\.", r"2\. Certification|2\.1\.", max_chars=900)
    m=re.search(r"1\.1\.1\.\s*([^\n]+)\s+Vessel Type", b, re.I)
    if not m:
        m=re.search(r"Vessel Type\s+([^\n]+)", b, re.I)
    if m:
        vt=clean(m.group(1))
        vt=re.sub(r"\b(Yes|No)\b.*$", "", vt, flags=re.I).strip()
        setf(f,"vessel_type",vt,"PIQ",b)
    if re.search(r"Annex II.*?\bYes\b",b,re.I|re.S): setf(f,"annex_ii_carriage","Yes","PIQ",b)

    # Superintendent sections: use block_between so decimal numbers like 59.00 do not stop extraction.
    b=block_between(text, r"2\.2\.1001\.", r"Has a Marine Superintendent|2\.2\.1002\.", max_chars=1800); ds=dates_in(b)
    if ds:
        setf(f,"technical_superintendent_dates",", ".join(dstr(x) for x in ds),"PIQ",b)
    b=block_between(text, r"2\.2\.1002\.", r"2\.3\. Structural|2\.3\.3001", max_chars=1800); ds=dates_in(b)
    if ds: setf(f,"marine_superintendent_dates",", ".join(dstr(x) for x in ds),"PIQ",b)
    for qid,key in [("2.3.3001","cargo_tank_oldest_inspection"),("2.3.3002","ballast_tank_oldest_inspection"),("2.3.3003","void_oldest_inspection")]:
        b=qblock(text,qid,70); ds=dates_in(b)
        if ds: setf(f,key,dstr(ds[-1]),"PIQ",b)
        m=re.search(r"Required frequency.*?(\d+)\s+months",b,re.I)
        if m: setf(f,key+"_freq",m.group(1),"PIQ",b)
    b=qblock(text,"2.5.1002",80); yn=yn_from_block(b)
    if yn: setf(f,"equipment_retrofitted",yn,"PIQ",b)
    if b: setf(f,"equipment_retrofit_details",b,"PIQ",b,"MEDIUM")
    b=qblock(text,"2.5.1003",100); yn=yn_from_block(b)
    if yn: setf(f,"equipment_replaced",yn,"PIQ",b)
    b=qblock(text,"2.8.2",120)
    # Extract last row after 'Last'.
    m=re.search(r"Last\s+((?:\d{1,2}\s+[A-Za-z]+\s+\d{4}|[A-Za-z]+\s+\d{1,2},\s*\d{4}))\s+(.+?)\s+(?:Indian Ocean MoU|Tokyo MoU|US Coastguard|USCG|Paris MoU|Korean MoU|Black Sea MoU).*?\s(Yes|No)\s+(Yes|No)",b,re.I|re.S)
    if m:
        setf(f,"last_psc_date",dstr(dates_in(m.group(1))[0]),"PIQ",b)
        setf(f,"last_psc_port",clean(re.sub(r",.*","",m.group(2))),"PIQ",b)
        setf(f,"psc_detained",m.group(3),"PIQ",b)
    sec=re.search(r"5\.7\..*?(?:6\. |6\.|$)",text,re.I|re.S)
    if sec:
        bb=sec.group(0)
        positives=re.findall(r"5\.7\.\d+\..{0,180}?\bYes\b",bb,re.I|re.S)
        setf(f,"incident_other","Yes" if positives else "No","PIQ",bb)
        setf(f,"incident_pollution_grounding_collision","No" if not positives else "Yes","PIQ",bb)
    else:
        # If incident section not extracted, do not assume mismatch; leave blank.
        pass
    return f

# -----------------------------
# Q88 extraction
# -----------------------------
def extract_q88(text: str) -> Dict[str, Field]:
    f={}
    if not text: return f
    b=qblock(text,"1.2",30); m=re.search(r"name.*?\(?IMO number\)?.*?([A-Za-z0-9 '\-]+)\s*\((\d{7})\)",b,re.I)
    if m: setf(f,"vessel_name",m.group(1),"Q88",b); setf(f,"imo_number",m.group(2),"Q88",b)
    b=qblock(text,"1.5",25); m=re.search(r"Flag/Port of Registry\s+(.+?)/(.+)",b,re.I)
    if m: setf(f,"flag",m.group(1),"Q88",b); setf(f,"port_registry",m.group(2),"Q88",b)
    b=qblock(text,"1.8",40); m=re.search(r"IOPPC\).*?\s+(Other\s*\(.+?\)|Oil Tanker|Chemical Tanker|.+?)(?:\s+1\.8a|$)",b,re.I)
    if m: setf(f,"vessel_type",m.group(1),"Q88",b)
    b=qblock(text,"1.10",80)
    # First name line after 'IMO Number'
    m=re.search(r"IMO Number\s+(.+?)(?:\s+c/o|\s+\d+|\s+Tel:|$)",b,re.I)
    if m: setf(f,"registered_owner",m.group(1),"Q88",b)
    b=qblock(text,"1.14",80); m=re.search(r"P\s*&\s*I Club.*?\s+(NorthStandard Limited|Northstandard Limitied|.+?)(?:\s+If other|\s+1\.15|$)",b,re.I)
    if m: setf(f,"pni_club",m.group(1),"Q88",b)
    b=qblock(text,"1.18",25); m=re.search(r"Classification society\s+(.+?)(?:\s+1\.18a|$)",b,re.I)
    if m: setf(f,"class_society",m.group(1),"Q88",b)
    b=qblock(text,"1.19",90); m=re.search(r"Class notation\s+(.+?)(?:\s+1\.20|$)",b,re.I)
    if m: setf(f,"class_notation",m.group(1),"Q88",b)
    for qid,key in [("1.20","conditions_of_class"),("1.20a","memoranda_of_class")]:
        b=qblock(text,qid,35); yn=yn_from_block(b)
        if yn: setf(f,key,yn,"Q88",b)
    for qid,keys in [("1.23",("last_drydock",None)),("1.24",("next_drydock_due","next_annual_survey_due")),("1.25",("last_special_survey","next_special_survey_due")),("1.25a",("last_iws","next_iws_due"))]:
        b=qblock(text,qid,45); ds=dates_in(b)
        if ds:
            setf(f,keys[0],dstr(ds[0]),"Q88",b)
            if keys[1] and len(ds)>1: setf(f,keys[1],dstr(ds[1]),"Q88",b)
            elif keys[1] and len(ds)==1: setf(f,keys[1],"", "Q88", b)
        elif qid=="1.25a" and b:
            setf(f,"last_iws", clean(re.sub(r".*due:","",b,flags=re.I)), "Q88", b, "LOW")
    # Q88 certificates: extract by numbered blocks 2.x when possible.
    qcert={
        "2.1":"safety_equipment","2.2":"safety_radio","2.3":"safety_construction","2.4":"loadline","2.5":"iopp",
        "2.6":"issc","2.9":"smc","2.10":"doc","2.11":"uscg_coc","2.12":"clc_oil","2.13":"clc_bunker",
        "2.14":"wreck_removal","2.15":"cofr","2.16":"class_certificate","2.19":"cof_chem","2.23":"ship_sanitation"
    }
    for qid,key in qcert.items():
        b=qblock(text,qid,45); ds=dates_in(b)
        if ds:
            # Q88 order: issue, annual, intermediate, expiry; but if only two dates, second is expiry.
            issue=dstr(ds[0])
            if len(ds)>=4: annual,inter,expiry=dstr(ds[1]),dstr(ds[2]),dstr(ds[3])
            elif len(ds)==3: annual,inter,expiry=dstr(ds[1]),"",dstr(ds[2])
            elif len(ds)==2: annual,inter,expiry="","",dstr(ds[1])
            else: annual,inter,expiry="","",""
            add_cert_field(f,key,"cert",issue,expiry,"Q88",b,annual,inter)
    # Q88 environmental section often uses 10.10 rather than HVPQ numbering.
    b=qblock(text,"10.10",70) or block_between(text,r"Does the vessel have a CII Rating number",max_chars=900)
    if b:
        m=re.search(r"Yes,\s*([A-E])", b, re.I) or re.search(r"CII rating[: ]+([A-E])", b, re.I)
        if m: setf(f,"cii_rating",m.group(1),"Q88",b)
        m=re.search(r"verified by Class, 3rd Party or Owner\??\s*(Class|Owner|3rd Party)", b, re.I)
        if m: setf(f,"cii_verified_by",m.group(1),"Q88",b)
    # Search CII in whole text as a fallback. If the verification question is present but blank, rating is still captured and basis left blank.
    m=re.search(r"CII Rating number.*?Yes,\s*([A-E]).{0,250}?verified by Class, 3rd Party or Owner\??\s*(Class|Owner|3rd Party)?",text,re.I|re.S)
    if m:
        setf(f,"cii_rating",m.group(1),"Q88",m.group(0))
        if m.group(2): setf(f,"cii_verified_by",m.group(2),"Q88",m.group(0))
    return f

# -----------------------------
# Class Status extraction
# -----------------------------
def extract_class(text: str) -> Dict[str, Field]:
    f={}
    if not text: return f
    if re.search(r"KOREAN REGISTER",text,re.I): setf(f,"class_society","Korean Register","Class Status",text[:400])
    elif re.search(r"NIPPON KAIJI KYOKAI|NK-SHIPS",text,re.I): setf(f,"class_society","Nippon Kaiji Kyokai","Class Status",text[:400])
    m=re.search(r"Ship Name\s+([A-Z0-9 '\-]+).*?IMO No\.\s*(\d{7})",text,re.I|re.S)
    if m: setf(f,"vessel_name",m.group(1),"Class Status",m.group(0)); setf(f,"imo_number",m.group(2),"Class Status",m.group(0))
    m=re.search(r"Flag\s+([A-Z ]+?)\s+IMO Number",text,re.I|re.S)
    if m: setf(f,"flag",m.group(1),"Class Status",m.group(0))
    m=re.search(r"Class Notation\s+(.+?)\s+Owner",text,re.I|re.S)
    if m: setf(f,"class_notation",m.group(1),"Class Status",m.group(0))
    m=re.search(r"Type of Ship.*?:\s*([^\n]+)",text,re.I)
    if m: setf(f,"vessel_type",m.group(1),"Class Status",m.group(0))
    m=re.search(r"\nOwner\s+([^\n]+?)\s*\nManager", text, re.I)
    if not m:
        m=re.search(r"\nOwner\s+([A-Z0-9 .,&'\-]+)", text, re.I)
    if m: setf(f,"registered_owner",m.group(1),"Class Status",m.group(0))
    # Class KR certificates: segment by labels, first date is issue, second date expiry.
    all_aliases=[a for v in CERT_ALIASES.values() for a in v]
    for key,aliases in CERT_ALIASES.items():
        seg=cert_segment(text, aliases, all_aliases, 550)
        if not seg: continue
        ds=dates_in(seg)
        if ds:
            issue=dstr(ds[0])
            # KR VGP row normally has issue only and a dash for expiry; do not convert next document row into expiry.
            expiry="" if key=="vgp" else (dstr(ds[1]) if len(ds)>1 else "")
            add_cert_field(f,key,"cert",issue,expiry,"Class Status",seg)
    # KR survey rows from Class Surveys section.
    class_survey_section=block_between(text,r"Class Surveys",r"Cargo Handling Appliances|Statutory Surveys",max_chars=1800)
    def survey_row(name):
        seg=block_between(class_survey_section, re.escape(name), r"(?:Special Survey|Intermediate Survey|Annual Survey|Docking Survey|No\.1 Propeller|No\.1 Aux|Due :)", max_chars=450)
        if not seg:
            m=re.search(re.escape(name)+r"(.{0,300})",class_survey_section,re.I|re.S)
            seg=name+(m.group(1) if m else "")
        return seg
    for name,lastkey,nextkey in [("Special Survey","last_special_survey","next_special_survey_due"),("Annual Survey","last_annual_survey","next_annual_survey_due"),("Intermediate Survey","last_intermediate_survey","next_intermediate_survey_due"),("Docking Survey","last_drydock","next_drydock_due")]:
        seg=survey_row(name); ds=dates_in(seg)
        if len(ds)>=1: setf(f,lastkey,dstr(ds[0]),"Class Status",seg)
        if len(ds)>=2: setf(f,nextkey,dstr(ds[1]),"Class Status",seg)
    # Conditions/Notes
    m=re.search(r"Condition of Class / Statutory Condition\s+Conditions\s+(.+?)(?:\n|14-May|\d{1,2}-[A-Za-z]+-\d{4})",text,re.I|re.S)
    if m:
        v=m.group(1); setf(f,"conditions_of_class","No" if re.search(r"Nil|None|No",v,re.I) else v,"Class Status",m.group(0))
    m=re.search(r"Actionable Note\s+Notes\s+(.+?)(?:\n|14-May|\d{1,2}-[A-Za-z]+-\d{4})",text,re.I|re.S)
    if m:
        v=m.group(1); setf(f,"memoranda_of_class","No" if re.search(r"Nil|None|No",v,re.I) else v,"Class Status",m.group(0))
    return f

# -----------------------------
# Rules
# -----------------------------
def values(f: Dict[str,Field], key: str) -> str:
    return f.get(key, Field()).value

def evidence(*fields: Field) -> str:
    return "\n---\n".join([f"{x.source}: {x.evidence}" for x in fields if x and x.evidence])

def add_finding(rows, area, check, status, risk, reason, action, hvpq="", piq="", cls="", q88="", ev=""):
    rows.append(Finding(area,check,status,risk,hvpq,piq,cls,q88,reason,action,ev))

def compare_field(rows, key, area, label, hvpq, other, other_name, risk="HIGH", manual_for_type=False):
    h=hvpq.get(key,Field()); o=other.get(key,Field())
    if not h.value or not o.value: return
    if equivalent(h.value,o.value,key): return
    if manual_for_type:
        add_finding(rows,area,label+f": HVPQ vs {other_name}","MANUAL CHECK","MEDIUM",f"Wording differs but may be acceptable depending certificate/Form B basis.","Verify wording against IOPP/COF/class and make HVPQ/PIQ/Q88 consistent.",hvpq=h.value, **({"piq":o.value} if other_name=="PIQ" else {"q88":o.value} if other_name=="Q88" else {"cls":o.value}), ev=evidence(h,o))
    else:
        add_finding(rows,area,label+f": HVPQ vs {other_name}","MISMATCH",risk,"Extracted values differ after normalization.","Verify latest source and correct HVPQ/Q88/PIQ as applicable.",hvpq=h.value, **({"piq":o.value} if other_name=="PIQ" else {"q88":o.value} if other_name=="Q88" else {"cls":o.value}), ev=evidence(h,o))

def run_rules(hvpq,piq,cls,q88,asof: date) -> List[Finding]:
    rows=[]
    # 1 superintendent strict rules
    for key,label,max_months in [("technical_superintendent_dates","Technical Superintendent visit gap",7),("marine_superintendent_dates","Marine Superintendent visit gap",12)]:
        fld=piq.get(key,Field())
        ds=dates_in(fld.value)
        # Dates appear as from/to pairs: last-from,last-to,second-from,second-to...
        pairs=[]
        for i in range(0,len(ds)-1,2):
            pairs.append((ds[i],ds[i+1]))
        # sort chronological by from date
        pairs=sorted(pairs,key=lambda x:x[0])
        for (prev_from,prev_to),(next_from,next_to) in zip(pairs,pairs[1:]):
            gap_days=(next_from-prev_to).days
            allowed=(prev_to+relativedelta(months=max_months))
            if next_from>allowed:
                add_finding(rows,"Management Oversight",label,"MISMATCH","CRITICAL",f"Gap from {prev_to} to {next_from} is {gap_days} days, exceeding strict {max_months}-month rule.","Office/vessel to provide explanation and schedule/record corrective inspection where required.",piq=fld.value,ev=fld.evidence)
    # 2 vessel type as manual only, not high mismatch
    compare_field(rows,"vessel_type","General","Vessel type wording",hvpq,piq,"PIQ",manual_for_type=True)
    compare_field(rows,"vessel_type","General","Vessel type wording",hvpq,q88,"Q88",manual_for_type=True)
    # 3 owner / P&I
    compare_field(rows,"registered_owner","Ownership","Registered owner",hvpq,q88,"Q88",risk="MEDIUM")
    compare_field(rows,"registered_owner","Ownership","Registered owner",hvpq,cls,"CLASS",risk="MEDIUM")
    compare_field(rows,"pni_club","Insurance","P&I club spelling / naming",hvpq,q88,"Q88",risk="MEDIUM")
    if "limitied" in values(hvpq,"pni_club").lower():
        add_finding(rows,"Insurance","P&I club spelling appears incorrect in HVPQ","MANUAL CHECK","MEDIUM","HVPQ appears to contain a spelling error in the P&I club name.","Correct spelling/name style in HVPQ after checking P&I certificate/Q88.",hvpq=values(hvpq,"pni_club"),q88=values(q88,"pni_club"),ev=evidence(hvpq.get("pni_club"),q88.get("pni_club")))
    # 4 Certificate comparisons targeted
    cert_labels={
        "cofr":"COFR", "iopp":"IOPPC", "vgp":"Vessel General Permit", "doc":"DOC", "class_certificate":"Class Certificate",
        "safety_equipment":"Safety Equipment", "safety_radio":"Safety Radio", "safety_construction":"Safety Construction", "loadline":"Loadline",
        "cof_chem":"Certificate of Fitness - Chemical", "uscg_coc":"USCG COC", "smc":"SMC", "issc":"ISSC", "ibwmc":"IBWMC",
        "clc_oil":"CLC Oil", "clc_bunker":"CLC Bunker", "wreck_removal":"Wreck Removal"
    }
    for ckey,lab in cert_labels.items():
        # compare expiry against Q88/Class
        for src,name in [(q88,"Q88"),(cls,"CLASS")]:
            if ckey=="uscg_coc" and name=="CLASS":
                continue
            hk=f"cert_{ckey}_expiry"; ok=f"cert_{ckey}_expiry"
            if values(hvpq,hk) and values(src,ok) and not equivalent(values(hvpq,hk),values(src,ok),hk):
                risk="CRITICAL" if ckey=="cofr" and dates_in(values(hvpq,hk)) and dates_in(values(hvpq,hk))[0] < asof else "HIGH"
                add_finding(rows,"Certificates",f"{lab} expiry: HVPQ vs {name}","MISMATCH",risk,"Certificate expiry differs; HVPQ may be outdated.","Verify latest certificate/Class/Q88 and correct HVPQ.",hvpq=values(hvpq,hk), q88=values(src,ok) if name=="Q88" else "", cls=values(src,ok) if name=="CLASS" else "", ev=evidence(hvpq.get(hk),src.get(ok)))
        # compare issue date against Class for selected statutory certificates
        if ckey in {"iopp","vgp","class_certificate","safety_equipment","safety_radio","safety_construction","loadline","ibwmc","cof_chem"}:
            hk=f"cert_{ckey}_issue"; ck=f"cert_{ckey}_issue"
            if values(hvpq,hk) and values(cls,ck) and not equivalent(values(hvpq,hk),values(cls,ck),hk):
                add_finding(rows,"Certificates",f"{lab} issue date: HVPQ vs CLASS","MISMATCH","HIGH","Certificate issue date differs from Class Status. This can mean HVPQ is not updated after reissue/renewal.","Verify latest certificate issue date and update HVPQ if Class Status is current.",hvpq=values(hvpq,hk),cls=values(cls,ck),ev=evidence(hvpq.get(hk),cls.get(ck)))
    # explicit expiry overdue check for HVPQ certs
    for key,field in hvpq.items():
        if key.startswith("cert_") and key.endswith("_expiry"):
            ds=dates_in(field.value)
            if ds and ds[0] < asof:
                add_finding(rows,"Certificates",key.replace("cert_","").replace("_expiry","").upper()+" expired in HVPQ","MISMATCH","CRITICAL",f"HVPQ shows expiry {ds[0]}, before document/as-of date {asof}.","Update HVPQ with valid certificate or confirm certificate status.",hvpq=field.value,ev=field.evidence)
    # DOC date sequence sanity
    if values(hvpq,"cert_doc_issue") and values(hvpq,"cert_doc_annual"):
        di=dates_in(values(hvpq,"cert_doc_issue")); da=dates_in(values(hvpq,"cert_doc_annual"))
        if di and da and da[0] < di[0]:
            add_finding(rows,"Certificates","DOC endorsement date appears inconsistent","MANUAL CHECK","HIGH","HVPQ DOC annual/endorsement date appears earlier than DOC issue date.","Verify latest DOC and endorsement details; correct HVPQ if table entry is stale/wrong.",hvpq=f"Issue {values(hvpq,'cert_doc_issue')}; annual {values(hvpq,'cert_doc_annual')}",ev=evidence(hvpq.get('cert_doc_issue'),hvpq.get('cert_doc_annual')))
    # IWS next due missing
    if values(hvpq,"last_iws") and not values(hvpq,"next_iws_due"):
        add_finding(rows,"Class / Survey","IWS next due blank in HVPQ","MANUAL CHECK","HIGH","HVPQ declares last IWS but next IWS due is blank.","Verify if IWS remains applicable after renewal/drydock and update HVPQ/Q88 if required.",hvpq=values(hvpq,"last_iws"),ev=evidence(hvpq.get('last_iws')))
    if values(q88,"last_iws") and not values(q88,"next_iws_due"):
        add_finding(rows,"Class / Survey","IWS next due blank in Q88","MANUAL CHECK","HIGH","Q88 declares last IWS but next IWS due is blank.","Verify and update Q88 if applicable.",q88=values(q88,"last_iws"),ev=evidence(q88.get('last_iws')))
    # Surveys compare: do not compare last_iws vs class docking blindly; only due dates where relevant.
    for key,label in [("next_drydock_due","Next drydock/docking survey due"),("next_special_survey_due","Next special survey due")]:
        compare_field(rows,key,"Class / Survey",label,hvpq,cls,"CLASS",risk="HIGH")
        compare_field(rows,key,"Class / Survey",label,hvpq,q88,"Q88",risk="HIGH")
    # CII verification basis
    if values(hvpq,"cii_rating") and values(q88,"cii_rating") and values(hvpq,"cii_rating") == values(q88,"cii_rating") and values(hvpq,"cii_verified_by") and not values(q88,"cii_verified_by"):
        add_finding(rows,"Environmental","Q88 CII verification basis blank","MANUAL CHECK","MEDIUM","Q88 appears to carry CII rating but verification basis was not extracted/populated while HVPQ has verification basis.","Verify Q88 CII verification basis and update if blank.",hvpq=f"{values(hvpq,'cii_rating')} / {values(hvpq,'cii_verified_by')}",q88=values(q88,"cii_rating"),ev=evidence(hvpq.get('cii_verified_by'),q88.get('cii_rating')))
    # Incidents nil confirmation
    hinc=values(hvpq,"incident_other") or values(hvpq,"incident_pollution_grounding_collision")
    pinc=values(piq,"incident_other")
    if (not hinc or norm_compare(hinc)=="no") and (not pinc or norm_compare(pinc)=="no"):
        add_finding(rows,"Incidents","No incidents declared","MANUAL CHECK","MEDIUM","No incidents appear declared in HVPQ/PIQ from extracted sections. Positive nil confirmation is still required.","Ship/office to confirm no reportable machinery, navigation, mooring, pollution, security, injury or operational incidents occurred in the last 12 months.",hvpq=hinc,piq=pinc)
    elif hinc and pinc and not equivalent(hinc,pinc,"incident"):
        add_finding(rows,"Incidents","Incident declaration HVPQ vs PIQ","MISMATCH","HIGH","HVPQ and PIQ incident declaration differs.","Align incident declaration after office/vessel verification.",hvpq=hinc,piq=pinc,ev=evidence(hvpq.get('incident_other'),piq.get('incident_other')))
    # PIQ tank due checks
    for key,label in [("cargo_tank_oldest_inspection","Cargo/slop tank inspection sequence"),("ballast_tank_oldest_inspection","Ballast tank inspection sequence"),("void_oldest_inspection","Void space inspection sequence")]:
        old=values(piq,key); freq=values(piq,key+"_freq") or "12"
        if old:
            ds=dates_in(old)
            if ds:
                due=ds[0]+relativedelta(months=int(float(freq)))
                risk="HIGH" if due < asof else "MEDIUM"
                status="MISMATCH" if due < asof else "MANUAL CHECK"
                add_finding(rows,"Tank Inspection",label+" due calculation",status,risk,f"Oldest date {ds[0]} with {freq}-month frequency gives due {due}.","Vessel to verify current inspection sequence and update PIQ if any tank/void inspection is outside sequence.",piq=f"Oldest {old}; freq {freq}; due {due}",ev=evidence(piq.get(key)))
    # Publications targeted check if HVPQ has stale-looking values
    if "publications_block" in hvpq:
        ev=hvpq["publications_block"].evidence
        stale=[]
        for item in ["COLREGS", "USCG Regulations", "Oil Transfer Procedures", "International Medical Guide", "MFAG"]:
            if re.search(item, ev, re.I): stale.append(item)
        if stale:
            add_finding(rows,"Publications","Publication editions targeted verification","MANUAL CHECK","MEDIUM","HVPQ publication section should be verified against latest onboard publication list/SMS controlled list.","Ship to confirm listed editions are current and update HVPQ where outdated/blank/NA.",hvpq="; ".join(stale),ev=ev)
    # MOC/retrofit targeted
    if values(piq,"equipment_retrofitted")=="Yes" or values(piq,"equipment_replaced")=="Yes":
        add_finding(rows,"Management of Change","PIQ retrofit/replacement declared - verify HVPQ/class/cert updates","MANUAL CHECK","MEDIUM","PIQ declares equipment retrofit/replacement. Associated HVPQ fields, certificates and class status should reflect latest condition.","Ship/office to verify MOC evidence, certificate reissue, Class survey and HVPQ update.",piq=values(piq,"equipment_retrofit_details"),ev=evidence(piq.get('equipment_retrofit_details')))
    # De-duplicate similar rows
    seen=set(); unique=[]
    for r in rows:
        k=(r.area,r.check,r.hvpq_value,r.piq_value,r.class_value,r.q88_value)
        if k not in seen:
            seen.add(k); unique.append(r)
    risk_order={"CRITICAL":0,"HIGH":1,"MEDIUM":2,"LOW":3}
    unique.sort(key=lambda r:(risk_order.get(r.risk,9),r.area,r.check))
    return unique

# -----------------------------
# Observation/checklist
# -----------------------------
def parse_obs_excel(uploaded) -> pd.DataFrame:
    if uploaded is None: return pd.DataFrame()
    try:
        df=pd.read_excel(uploaded)
    except Exception:
        return pd.DataFrame()
    if df.empty: return df
    df.columns=[str(c).strip() for c in df.columns]
    joined=df.fillna("").astype(str).apply(lambda r:" ".join(r.tolist()),axis=1)
    df["row_text"]=joined
    def fam(x):
        y=x.lower()
        if any(w in y for w in ["incident","injury","ground","collision","allision"]): return "Incident declaration"
        if any(w in y for w in ["certificate","expiry","issued","endorsement","cofr"]): return "Certificate/date accuracy"
        if any(w in y for w in ["class","condition","memoranda","dry dock","iws","survey"]): return "Class/survey"
        if any(w in y for w in ["mooring","brake","rope","tail","winch","fairlead","chock"]): return "Mooring"
        if any(w in y for w in ["tank","coating","void","ballast"]): return "Tank/structural"
        if any(w in y for w in ["foam","fire","lifeboat","rescue"]): return "Firefighting/LSA"
        if any(w in y for w in ["piping","overboard","sea chest","scupper","pressure"]): return "Pollution/cargo"
        if any(w in y for w in ["diagram","manifold","arrangement"]): return "Diagrams"
        return "Other HVPQ accuracy"
    df["family"]=joined.map(fam)
    return df

def checklist(obs: pd.DataFrame) -> pd.DataFrame:
    rows=[
        ["PIQ/HVPQ","Incident declaration","Confirm nil incidents or declare all reportable machinery/navigation/mooring/pollution/security/injury/operational incidents in last 12 months."],
        ["Class Status","Certificates","Check COFR, IOPP, VGP, Class, SEC/SRC/SCC/Loadline/COF/SMC/ISSC issue and expiry dates against latest certificates/Class Status."],
        ["Class Status","COC/MOC/Notes","Confirm Conditions of Class, statutory conditions, actionable notes, memoranda/notes and dispensations are declared correctly."],
        ["PIQ","Superintendent visits","Check Technical Superintendent gap <=7 months and Marine Superintendent gap <=12 months; any exceedance to be explained/actioned."],
        ["PIQ","Tank inspections","Check cargo/slop, ballast and void inspection sequence dates against tank inspection records."],
        ["HVPQ/Q88","Mooring","Verify brake test date, BHC/rendering load, split drum, mooring rope/tail certificates, end-for-end/discard and diagrams."],
        ["HVPQ/Q88","Cargo/pollution","Verify cargo/bunker pressure test, overboard blanks/testing, sea chest/scupper and cargo system declarations."],
        ["HVPQ/Q88","Firefighting/LSA","Verify foam type/test date, fixed systems, sample locker systems, rescue boat/davit and LSA certificates."],
        ["HVPQ","Publications","Verify publication editions against onboard controlled publication list."],
    ]
    if obs is not None and not obs.empty and "family" in obs:
        for fam,cnt in obs["family"].value_counts().items():
            rows.append(["Observation library",fam,f"Historical observations contain {cnt} item(s) in this family. Include in ship verification."])
    return pd.DataFrame(rows,columns=["Source","Area","Targeted check"])

def fields_df(srcname, fields):
    return pd.DataFrame([{"source":srcname,"field":k,"value":v.value,"confidence":v.confidence,"evidence":v.evidence[:400]} for k,v in sorted(fields.items())])

# -----------------------------
# Streamlit
# -----------------------------
def main():
    st.set_page_config(page_title="HVPQ / PIQ Checker v8", layout="wide")
    st.title("HVPQ / PIQ Vetting Observation Checker v8")
    st.caption("Extraction-first. Source-aware. No broad class/Q88 fuzzy mismatches. Designed to produce targeted ship/office verification checks.")
    with st.sidebar:
        st.header("Upload documents")
        hvpq_file=st.file_uploader("HVPQ PDF",type=["pdf"])
        piq_file=st.file_uploader("PIQ PDF",type=["pdf"])
        class_file=st.file_uploader("Class Status PDF",type=["pdf"])
        q88_file=st.file_uploader("Q88 PDF",type=["pdf"])
        obs_file=st.file_uploader("Observation library Excel (optional)",type=["xlsx","xls"])
        asof=st.date_input("As-of / review date", value=date.today())
        show_evidence=st.checkbox("Show evidence column",value=False)
    hvpq_txt=pdf_text(hvpq_file)
    piq_txt=pdf_text(piq_file)
    cls_txt=pdf_text(class_file)
    q88_txt=pdf_text(q88_file)
    hvpq=extract_hvpq(hvpq_txt)
    piq=extract_piq(piq_txt)
    cls=extract_class(cls_txt)
    q88=extract_q88(q88_txt)
    # Use document date as as-of when available and user didn't change? Keep user input but default today. Findings use selected.
    findings=run_rules(hvpq,piq,cls,q88,asof)
    df=pd.DataFrame([asdict(x) for x in findings])
    if not show_evidence and not df.empty:
        df_show=df.drop(columns=["evidence"])
    else:
        df_show=df
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Actionable rows",len(df))
    c2.metric("Critical",int((df["risk"]=="CRITICAL").sum()) if not df.empty else 0)
    c3.metric("High",int((df["risk"]=="HIGH").sum()) if not df.empty else 0)
    c4.metric("Manual/Medium",int((df["status"]=="MANUAL CHECK").sum()) if not df.empty else 0)
    tabs=st.tabs(["Findings","Extracted fields","Ship checklist","Observation library","Debug text"])
    with tabs[0]:
        st.subheader("Actionable mismatch / manual-check register")
        st.dataframe(df_show,use_container_width=True,height=520)
        if not df.empty:
            bio=io.BytesIO()
            with pd.ExcelWriter(bio,engine="openpyxl") as writer:
                df.to_excel(writer,index=False,sheet_name="Findings")
                pd.concat([fields_df("HVPQ",hvpq),fields_df("PIQ",piq),fields_df("Class",cls),fields_df("Q88",q88)],ignore_index=True).to_excel(writer,index=False,sheet_name="Extracted fields")
                checklist(parse_obs_excel(obs_file)).to_excel(writer,index=False,sheet_name="Ship checklist")
            st.download_button("Download Excel register",bio.getvalue(),"hvpq_piq_v8_register.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tabs[1]:
        st.dataframe(pd.concat([fields_df("HVPQ",hvpq),fields_df("PIQ",piq),fields_df("Class",cls),fields_df("Q88",q88)],ignore_index=True),use_container_width=True,height=600)
    with tabs[2]:
        st.dataframe(checklist(parse_obs_excel(obs_file)),use_container_width=True,height=600)
    with tabs[3]:
        obs=parse_obs_excel(obs_file)
        if obs.empty: st.info("No observation Excel uploaded or no rows read.")
        else: st.dataframe(obs,use_container_width=True,height=600)
    with tabs[4]:
        st.write("Text lengths", {"hvpq":len(hvpq_txt),"piq":len(piq_txt),"class":len(cls_txt),"q88":len(q88_txt)})
        with st.expander("HVPQ text sample"): st.text(hvpq_txt[:5000])
        with st.expander("Class text sample"): st.text(cls_txt[:5000])
        with st.expander("Q88 text sample"): st.text(q88_txt[:5000])

if __name__ == "__main__":
    main()
