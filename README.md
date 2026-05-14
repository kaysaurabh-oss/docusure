# HVPQ / PIQ Vetting Observation Checker

A Streamlit app for objective, targeted HVPQ/PIQ checks before SIRE/vetting inspections.

## What it checks
- HVPQ vs PIQ consistency
- Incident declaration mismatch / blank declarations
- PIQ general accuracy against HVPQ
- HVPQ certificate/survey dates against Class Status
- Optional Q88 cross-check against HVPQ
- Observation-library driven targeted checks
- Simple export register for ship verification

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Important design choice
This is a rule-based audit assistant. It does not certify the HVPQ as correct. It produces targeted checks and manual verification items that can be exported and sent to the ship.
