# HVPQ / PIQ Vetting Observation Checker v4

## What this version changes

- Uses a mismatch-first workflow: no PASS rows shown by default.
- Uses the historical HVPQ / Incident Observation Excel files as a targeted check library.
- Adds optional HVPQ XML ingestion.
- Adds support for Class Status and Q88 / vessel particulars in different formats.
- Uses loose semantic comparison for equivalent labels/values such as Product Carrier vs Products/Chemical Tanker.
- Flags incident declarations simply: HVPQ vs PIQ mismatch, or no incident declared requiring positive confirmation.
- Generates an exportable Excel register that can be sent to vessel/office.

## Important limitation

OCIMF response XML normally contains control GUIDs and response values, not human-readable question labels. For question-level XML accuracy, upload a control mapping file with columns like:

```text
ctrl,qid,label
5ACB5920-94A6-46BB-AD3D-D233048F011F,1.1.1,Date this HVPQ document completed
```

Without this map, the app uses HVPQ PDF for question labels and XML for document metadata / identity.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Upload order

1. HVPQ PDF
2. HVPQ XML if available
3. PIQ PDF
4. Class Status / Certificate Status PDF/XLSX/TXT
5. Q88 if available
6. HVPQ Observation Excel
7. Incident Observation Excel

## Output

- Findings tab: mismatch and manual-check register
- Targeted ship checklist tab: ship-facing simplified checks
- Observation library tab: what the observation files contributed
- Extracted fields tab: debug extraction
- XML diagnostic tab: whether XML is usable directly or needs control mapping
