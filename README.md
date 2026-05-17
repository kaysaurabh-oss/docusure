# HVPQ / PIQ / Q88 / Class Status Checker v17

Extraction-first Streamlit checker for SIRE/HVPQ preparation.

## What changed in v17

- HVPQ remains the main document to correct.
- Class Status is used as the authority for certificate/survey dates and Conditions/Memoranda/dispensations.
- Q88 is used only as value-add cross-check, not as an authority by itself.
- PIQ is extracted and checked as an operational declaration document.
- Tank coating checks are now table-aware: the app checks **Date of last coating inspection** and does not use the original coating date.
- Observation-history questions are not converted into bulk manual checks. They are used only to strengthen reasons where a real HVPQ issue, mismatch, stale date or uncertainty is found.
- Output is split into three user-facing registers: HVPQ Checks, Q88 Value Add, PIQ Checks.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Optional local LLM extraction assist

Install Ollama and pull a model, for example:

```bash
ollama pull qwen2.5:14b
```

Then enable the local LLM extraction option in the sidebar.
