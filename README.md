# DAQ Tool

Phone-first weld-by-weld data logger for pole welding sessions. One self-contained HTML page:
session header, per-weld goal/method/flats/result, one-tap start/stop timing (UTC), accidental-restart
segments, before/after photos, Backup/Restore, and CSV/TSV/ZIP export that matches `DAQTool-Template.xlsx`.

All entered data stays in the phone's browser storage. Nothing is sent anywhere unless the user exports it.

- `index.html` - the app (open in Safari, Add to Home Screen)
- `DAQTool-Template.xlsx` - Excel workbook matching the export column-for-column
- `make_template.py` - regenerates the workbook (`python make_template.py`)
- `build_workbook.py` - turns a Share-bundle ZIP into a workbook with the photos embedded per row (`python build_workbook.py bundle.zip`)
