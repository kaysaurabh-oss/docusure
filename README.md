# HVPQ / PIQ Vetting Observation Checker v7

Extraction-first Streamlit app.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Design changes in v7

- HVPQ certificate table parsing uses row segmentation and takes expiry as the second date.
- Q88 certificate parsing uses Q88 table order and takes expiry as the last date.
- Class Status parsing is restricted to ClassNK style certificate expiry, survey status, conditions/notes and vessel particulars.
- PIQ extraction is restricted to vessel type, PSC, superintendent visits, tank inspection dates, MOC/retrofit and incident declaration section.
- No mismatch is produced when a field is not extracted. Missing extraction is visible in the Extracted Fields tab.
- Observation Excel is used to generate targeted ship-checklist categories, not as direct evidence.
