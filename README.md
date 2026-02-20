# Excel Columns Merger

Simple Streamlit app to merge two Excel sheets, optionally run similarity matching, and download one output Excel file.

## What This App Does

- Upload 2 Excel files (left and right).
- Pick which columns you want in the output.
- Optional: pick similarity columns to enable matching mode.
- If similarity columns are not selected, the app runs merge mode only.
- Filter result rows (`All`, `Matched only`, `No match only`) when matching mode is enabled.
- Save and load presets so repeated jobs are faster.
- Download result as `.xlsx`.

## Requirements

- Windows, macOS, or Linux
- Python 3.10+ recommended

## Install

Open terminal in the project folder, then run:

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

macOS/Linux:

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

Recommended:

```bash
python -m streamlit run app.py
```

Windows shortcuts:

- `run_app.bat`
- `run_app.ps1`

If terminal says `streamlit is not recognized`, use:

```bash
python -m streamlit run app.py
```

## How To Use

1. Upload left and right Excel files.
2. Select the sheet from each file.
3. Choose output columns from both sides.
4. Optional: choose similarity columns on both sides.
   - Selected on both sides: matching mode (`Similarity_Score`, `Match_Status`) is enabled.
   - Not selected: merge mode only (no similarity columns in output).
5. Click `Run Merge`.
6. Review result and download the merged Excel file.

## Presets

- Use `Preset name` + `Save preset` to save current settings.
- Use `Saved Preset` + `Apply` to restore settings.
- Use `Delete` to remove a preset.

Presets are stored in:

`presets/settings_presets.json`

## Notes

- Matching can be accelerated with `rapidfuzz` when enabled.
- In merge-only mode, rows are aligned by row order.
- App writes no source Excel changes; only generated output is downloaded.

## Troubleshooting

- `ModuleNotFoundError`: install packages again with `pip install -r requirements.txt`.
- `streamlit not recognized`: use `python -m streamlit run app.py`.
- If app does not refresh preset changes, click `Apply` again or rerun the app.
"# Excels-Columns-Merger" 
