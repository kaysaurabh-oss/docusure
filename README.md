# HVPQ / PIQ / Q88 / Class Status Checker v13

Extraction-first Streamlit app for pre-vetting HVPQ/PIQ/Q88/Class Status verification.

## Core workflow

- HVPQ is treated as the main document to correct.
- Class Status is treated as the authority only for certificate/survey dates, Conditions of Class, Memoranda and dispensations.
- Q88 is treated as value-add. Q88 mismatches/blanks are highlighted separately and should be verified against source evidence before correcting HVPQ.
- PIQ is checked for operational declarations such as superintendent visits, tank inspection cycles, incidents, MOC/retrofit and PSC.
- Observation Excel question numbers are extracted and converted into targeted HVPQ/vessel checks.
- If the app cannot reliably extract/check an item, it is flagged in Manual Confirmation so it is not silently missed.

## UI tabs

1. Review Dashboard - simple paragraph summary and top action items.
2. Vessel Register - clean action list suitable for vessel/office.
3. HVPQ / PIQ Issues - office correction register.
4. Manual Confirmation - data not reliably checked or requiring positive confirmation.
5. Q88 Value Add - Q88-specific mismatch/blank review.
6. Coverage - what was checked and what needs manual confirmation.
7. Advanced Review - optional extracted fields/raw text/debug.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Optional local LLM extraction assist through Ollama:

```bash
ollama pull qwen2.5:14b
ollama run qwen2.5:14b
```

Then enable the Ollama checkbox in the sidebar.
