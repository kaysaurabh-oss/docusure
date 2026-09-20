# DocuSure — SIRE & Chartering Readiness v22

DocuSure is an evidence-led Streamlit checker for tanker document packs. Upload all available PDFs, XML, Word, Excel or CSV files into one drop zone; the app first normalises each file into searchable text plus table/sheet rows, classifies each document, groups it by checksum-valid IMO and lets the user run the check needed from the same pack.

## Available checks

- **Data Accuracy** — HVPQ/PIQ/Q88/Class consistency, dates, blanks, logic and source evidence.
- **SIRE / Chartering Readiness** — prioritised concerns, key vessel facts and evidence gaps that may attract vetting or chartering questions.
- **STS Clearance Pre-Check** — choose **Single-vessel readiness** or **Two-vessel compatibility**. Single mode checks certificate expiry/short-term validity, P&I International Group membership, Class Status alignment, Conditions of Class, mooring renewal/service dates, one-year winch BRC, Q88 STS=Yes and last-ten-port sanctions cues. Two-vessel mode adds the compatibility fact screen. Both are explicitly pre-screens, not final operational clearance.
- **Certificate & Class Watch** — expired and short-validity certificates, Conditions/Memoranda of Class, dispensations and survey windows.
- **Mooring Readiness** — brake-test information, SDMBL/LDBF/TDBF and cautiously recognised rope/wire/tail rows.
- **Document Pack Completeness** — shows which evidence is present or missing for each use case.

## How the accuracy-first reader works

DocuSure no longer trusts one extraction route. In the default accuracy-first mode it builds a temporary evidence layer before running any rule:

1. **Native structure** — PDF text, Word paragraphs/tables, workbook sheets/rows and XML are read without flattening their structure.
2. **Layout-aware Markdown** — PyMuPDF4LLM performs page-chunked layout analysis for multi-column text and tables, while native PyMuPDF and Poppler provide independent readings; the strongest page representation is retained.
3. **Table channel** — detected PDF/Word/Excel tables remain separate row arrays as well as searchable text.
4. **OCR recovery** — sparse pages use Tesseract; remaining unreadable pages can pass through a temporary OCRmyPDF rotate/deskew copy when the host provides it.
5. **Optional local page-image AI** — a vision-capable Ollama model can transcribe critical/unreadable pages or every page into literal Markdown. It is never asked to decide compliance.
6. **Schema extraction and evidence gate** — AI candidates must match a value back to a source page. Certificate dates additionally require the correct issue/expiry/annual/intermediate label. Unsupported AI output is discarded.
7. **Decision rules** — only the verified evidence layer reaches the accuracy, SIRE/chartering, STS, certificate/class and mooring checks.

The UI can download this intermediate representation as a Markdown + JSON ZIP, including page/sheet boundaries, detected tables, field confidence and evidence excerpts. This is the most useful way to diagnose a bad result: first inspect what the engine actually read, then inspect the rule.

## What v22 changes

- One multi-file uploader replaces document-specific upload boxes.
- The normalisation layer accepts PDF/XML, `.doc`/`.docx`/`.docm`, `.xlsx`/`.xlsm`/`.xls` and CSV. PDFs retain page text and detected table rows; Word retains paragraphs/table rows; workbooks retain sheet names and row boundaries. Legacy `.doc` is converted through LibreOffice when available.
- Automatic classification covers HVPQ, PIQ, Q88, Class Status, individual certificates, P&I evidence, crew matrices, last-ten-port history, sanctions evidence, STS plans, STS/JPO/compatibility documents, PSC reports and LMP/MSMP documents.
- Files are grouped by a checksum-valid IMO. Cross-document conclusions are blocked if different vessels are mixed.
- Classification confidence and vessel-grouping confidence are shown separately. Uncertain files stay unassigned or are clearly marked provisional, with an optional session-only correction control.
- A joint STS/JPO document containing two checksum-valid IMO numbers is recognised as shared evidence for both vessels.
- Scanned/sparse PDFs trigger an OCR attempt when OCR is available; otherwise they are clearly marked as poor extraction rather than reported as blank documents.
- Accuracy-first PDF ingestion now compares layout-aware PyMuPDF4LLM Markdown with native PyMuPDF and Poppler reading order, retains the strongest page representation, and can invoke a rotate/deskew OCR fallback.
- Optional local-AI extraction reads multiple relevant page chunks instead of only the beginning of a document. Its values enter the canonical source only after an exact page-evidence match and remain lower priority than native/table evidence.
- If rule-based classification still returns Unknown, optional local AI may classify the document only when it supplies a literal evidence quote that can be matched back to the file.
- Optional local vision reading sends rendered page images to a configured Ollama endpoint for literal Markdown transcription. It can target critical/unreadable pages or every page, with a configurable page cap.
- A downloadable machine-readable evidence pack exposes the normalized Markdown, fields, tables, page/sheet provenance, reading channels and fallback notes.
- Ambiguous numeric dates are parsed day-first, consistent with common maritime document usage; ISO dates remain year-first.
- Certificate dates are role-aware: an expiry is accepted only when its label or a visible table header says expiry/valid-until. Issue, annual and intermediate dates are kept separate; impossible issue/expiry pairs withhold the expiry and create an extraction-quality action instead of silently using the wrong date.
- A PDF table channel is appended to the machine-readable text so the engine can use row/column context rather than only PDF reading order. All extracted fields retain filename and page/sheet provenance.
- Competing extraction candidates are scored before rules consume a value. IMO checksum, date validity, extraction method and local evidence affect the score.
- Output distinguishes **Document-supported discrepancy**, **Potential concern**, **Not verified**, and **In order from available evidence**.
- Every main output carries evidence confidence and an explanation of what the result means.
- Extracted values retain their source filename and page/sheet so conclusions can be traced back to evidence.
- Short-term certificate windows are separated into expired, 30-day, 90-day and configurable watch periods.
- Conditions/Memoranda/dispensations require explicit evidence. A Class Status heading alone is never treated as an open item.
- Q88 remains value-add only and is not treated as SIRE authority.
- The Excel register now contains key facts, readiness, certificate/class watch, mooring information, source registers, evidence and the recognised file inventory.

## Privacy model

The application code does not write uploaded documents to a DocuSure database or permanent document store. Processing and the file-derived cache remain in the active Streamlit session. PDF/OCR/legacy-Word conversions use self-deleting temporary directories. The UI includes a **Clear documents from this session** control. No external AI service is called by default.

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

Then enable the option in **Review settings**. A separate vision-capable model may be configured for page-image transcription. Local-AI extraction has lower confidence than deterministic/table-aware extraction and cannot by itself create a confirmed discrepancy. Every accepted AI value must match source-page evidence, and deterministic date-role validation remains the authority, so an AI model cannot turn an issue date into an expiry.

## Optional system readers

The Python packages in `requirements.txt` are sufficient for the core app. Accuracy improves when the host also provides:

- `pdftotext` (Poppler) for layout-preserving independent PDF text;
- Tesseract for page OCR;
- `ocrmypdf` for rotation/deskew recovery of difficult scans;
- LibreOffice for legacy `.doc` conversion.

Each reader is optional and fails visibly in **Machine-readable conversion evidence** without stopping the rest of the pack.
For Debian/Streamlit-style deployments, the supplied `packages.txt` requests these system readers; `ocrmypdf` is also included in `requirements.txt`.

## Important boundary

DocuSure is a document pre-screen. It does not replace the Master, operator, class, flag, insurer, terminal, STS service provider or charterer's formal review. In particular, the STS result never constitutes final operational clearance.
