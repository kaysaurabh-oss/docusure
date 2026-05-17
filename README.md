# HVPQ / PIQ / Q88 / Class Status Checker v18

Extraction-first Streamlit checker for SIRE/HVPQ preparation.

## Main v18 change

- Removed external upload controls for HVPQ observation library Excel, incident observation library Excel, and Comparison Rules TXT.
- Observation history and validation rules are now embedded as a machine-readable knowledge base inside the app.
- Observations are used only as priority signals. The app does not create a defect unless extracted HVPQ/PIQ/Q88/Class data is missing, stale, contradictory, expired, or logically doubtful.
- Observation priorities are sorted by repeat count, with highest-repeat HVPQ question numbers first.
- Tank coating logic remains table-aware: it checks the date of last coating inspection, not the original coating/application date.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Bundled files

- `app.py` - complete Streamlit app with embedded rule base.
- `embedded_knowledge_base.json` - same knowledge base exported separately for review/version control.
- `requirements.txt` - Python package requirements.
