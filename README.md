# HVPQ / PIQ / Q88 / Class Status Checker v15

Extraction-first Streamlit app for HVPQ/PIQ/Q88/Class Status review.

## Key workflow
- HVPQ is treated as the main document to correct.
- Class Status is used only as the reference/authority for certificate dates, survey dates, Conditions of Class, Memoranda and dispensations.
- Q88 is treated as value-add: differences are highlighted separately for review, not as automatic authority.
- PIQ is reviewed for operational declarations, superintendent intervals, PSC, tank inspection cycles, MOC/retrofit, incident declarations and selected rule checks.
- Uploaded observation Excel question numbers are no longer shown as a separate checklist if the app can already check the item. They are merged into HVPQ check reasons as “high repeat-observation area”.
- If the app cannot reliably extract/check an item, it flags it as Manual so the user knows it was not ignored.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Optional local LLM
Install Ollama and run a local model if needed:
```bash
ollama pull qwen2.5:14b
```
Then enable Local LLM extraction assist in the sidebar.

## Excel export sheets
- HVPQ Checks
- Q88 Value Add
- PIQ Checks

All sheets use wrapped text, wider columns and priority coloring.
