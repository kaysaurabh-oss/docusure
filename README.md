# HVPQ / PIQ / Q88 / Class Status Checker v14

Extraction-first Streamlit app for pre-vetting checks.

## Workflow

1. Upload HVPQ PDF (main correction target).
2. Upload PIQ PDF.
3. Upload Q88 PDF, if available (value-add cross-check only).
4. Upload Class Status PDF (authority/reference for certificate and survey dates, Conditions/Memoranda/dispensations).
5. Upload HVPQ / incident observation Excel sheets to generate repeat-observation question checks.
6. Run checks and download the Excel register.

## Output tabs in the app

- Office review summary
- HVPQ Checks
- Q88 Checks
- PIQ Checks
- Repeat Observation Questions
- Advanced Extraction

## Excel export

The Excel export intentionally has four sheets only:

- HVPQ Checks
- Q88 Checks
- PIQ Checks
- Repeat Obs Questions

The sheets use wrapped text and readable row heights for vessel/office circulation.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Optional local LLM assist requires Ollama:

```bash
ollama pull qwen2.5:14b
```

Then enable the Ollama checkbox in the sidebar.
