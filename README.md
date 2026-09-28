# Proteasome QC & Technical-Replicate Reviewer

Post-processing Streamlit app for Excel workbooks exported by the Proteasome ProSight Native Mass Matcher.

## Features
- Reads the segmented `Combined results` sheet directly.
- Manual QC dropdown: Not reviewed, Keep/confident, Interesting/investigate, Low intensity, Noise, Ambiguous, Exclude.
- Reviewer notes retained in the downloadable workbook.
- Compares two technical replicates using an adjustable mass tolerance (default ±1 Da).
- Optional same-LC-peak requirement.
- Keeps replicate-only observations as flags rather than automatically calling them noise.
- Exports `Reviewed masses`, `Replicate comparison`, `Qualified masses`, and `Recurring candidates` sheets.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```
