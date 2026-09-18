# DocuSure — SIRE & Chartering Readiness v20

DocuSure is an evidence-led Streamlit checker for tanker document packs. Upload all available PDFs/XML files into one drop zone; the app classifies each document, groups it by checksum-valid IMO and lets the user run the check needed from the same pack.

## Available checks

- **Data Accuracy** — HVPQ/PIQ/Q88/Class consistency, dates, blanks, logic and source evidence.
- **SIRE / Chartering Readiness** — prioritised concerns, key vessel facts and evidence gaps that may attract vetting or chartering questions.
- **STS Clearance Pre-Check** — two-vessel document readiness and side-by-side compatibility fact screen. This is explicitly a pre-screen, not final operational clearance.
- **Certificate & Class Watch** — expired and short-validity certificates, Conditions/Memoranda of Class, dispensations and survey windows.
- **Mooring Readiness** — brake-test information, SDMBL/LDBF/TDBF and cautiously recognised rope/wire/tail rows.
- **Document Pack Completeness** — shows which evidence is present or missing for each use case.

## What v20 changes

- One multi-file uploader replaces document-specific upload boxes.
- Automatic classification covers HVPQ, PIQ, Q88, Class Status, individual certificates, P&I evidence, STS plans, STS/JPO/compatibility documents, PSC reports and LMP/MSMP documents.
- Files are grouped by a checksum-valid IMO. Cross-document conclusions are blocked if different vessels are mixed.
- Classification confidence and vessel-grouping confidence are shown separately. Uncertain files stay unassigned or are clearly marked provisional, with an optional session-only correction control.
- A joint STS/JPO document containing two checksum-valid IMO numbers is recognised as shared evidence for both vessels.
- Scanned/sparse PDFs trigger an OCR attempt when OCR is available; otherwise they are clearly marked as poor extraction rather than reported as blank documents.
- Ambiguous numeric dates are parsed day-first, consistent with common maritime document usage; ISO dates remain year-first.
- Competing extraction candidates are scored before rules consume a value. IMO checksum, date validity, extraction method and local evidence affect the score.
- Output distinguishes **Document-supported discrepancy**, **Potential concern**, **Not verified**, and **In order from available evidence**.
- Every main output carries evidence confidence and an explanation of what the result means.
- Extracted values retain their source filename and page so conclusions can be traced back to evidence.
- Short-term certificate windows are separated into expired, 30-day, 90-day and configurable watch periods.
- Conditions/Memoranda/dispensations require explicit evidence. A Class Status heading alone is never treated as an open item.
- Q88 remains value-add only and is not treated as SIRE authority.
- The Excel register now contains key facts, readiness, certificate/class watch, mooring information, source registers, evidence and the recognised file inventory.

## Privacy model

The application code does not write uploaded documents to a DocuSure database or permanent document store. Processing and the file-derived cache remain in the active Streamlit session. The UI includes a **Clear documents from this session** control. No external AI service is called by default.

If the app is deployed through a third-party host, that host's infrastructure, proxy and logging policies still apply. Confirm those policies before uploading confidential material. The optional Ollama assist should only be pointed at an endpoint controlled by the organisation.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Verify the build

```bash
python -m py_compile app.py docusure_engine.py
python -m json.tool embedded_knowledge_base.json > /dev/null
python -m unittest discover -s tests -v
```

## Optional OCR

PyMuPDF attempts OCR only for pages with almost no searchable text. OCR requires Tesseract and English language data on the host. If unavailable, the app continues safely and marks the document as poor/sparse text.

## Optional local LLM assist

Install Ollama and pull a model, for example:

```bash
ollama pull qwen2.5:14b
```

Then enable the option in **Review settings**. Local-LLM extraction has lower confidence than deterministic/table-aware extraction and cannot by itself create a confirmed discrepancy.

## Important boundary

DocuSure is a document pre-screen. It does not replace the Master, operator, class, flag, insurer, terminal, STS service provider or charterer's formal review. In particular, the STS result never constitutes final operational clearance.
