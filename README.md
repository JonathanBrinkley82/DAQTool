# DAQ Tool

Phone-first weld-by-weld data logger for pole welding sessions. One self-contained HTML page:
session header, per-weld goal/method/flats/result, one-tap start/stop timing (UTC), accidental-restart
segments, before/after photos, Backup/Restore, and CSV/TSV/ZIP export that matches `DAQTool-Template.xlsx`.

All entered data stays in the phone's browser storage. Nothing is sent anywhere unless the user exports it.

**Back up before the URL changes.** Browser storage is tied to the page address. If this app moves
to a new URL (for example, a GitHub org transfer), every phone will open the new address with zero
sessions. Before any move: tap Backup on each phone and save the file, then Restore it at the new
address. The old data stays under the old URL and is not migrated automatically.

- `index.html` - the app (open in Safari, Add to Home Screen)
- `DAQTool-Template.xlsx` - Excel workbook matching the export column-for-column
- `make_template.py` - regenerates the workbook (`python make_template.py`)
- `build_workbook.py` - turns a Share-bundle ZIP into a workbook with the photos embedded per row (`python build_workbook.py bundle.zip`)
