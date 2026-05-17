# HVPQ / PIQ / Q88 / Class Status Checker v12

Extraction-first Streamlit app for HVPQ correction and vessel-ready verification.

## v12 changes
- Removed Office Summary from Excel export.
- Added Coverage Matrix so users know what was checked and what could not be reliably checked.
- Added Manual Checks sheet for extraction gaps/uncertain items.
- HVPQ is treated as the main correction document.
- Class Status is treated as the authoritative source for certificate, class and survey dates where available.
- Q88 mismatches are separated into a Q88 Value Add sheet and not mixed with authoritative Class Status findings.
- Observation Excel question numbers like `10.1.4`, `2.1.5`, `1.9.8` are extracted and checked against the HVPQ text.
- Vessel Action Checklist is clearer and formatted for sending to vessel/office.
- Excel export uses wrapped text, readable row heights and sensible columns.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Optional local LLM
```bash
ollama pull qwen2.5:14b
```
Enable local LLM extraction assist in the app sidebar.
