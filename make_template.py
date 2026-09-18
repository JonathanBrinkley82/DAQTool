"""Builds PoleWeldLog-Template.xlsx: a workbook whose Weld Log sheet matches the
app's CSV/TSV export column-for-column, with dropdowns, filters, and a runtime
check column. Run:  python make_template.py"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

COLS = [
    ("session_date", 12), ("shift", 8), ("session", 9), ("unit", 6),
    ("pole_side", 10), ("flange", 8), ("weld_side", 10), ("direction", 10),
    ("rotation", 9), ("setup_notes", 40), ("recorder_clock_reading", 20),
    ("phone_time_at_reading_utc", 24), ("std_wfs", 9), ("std_volts", 10),
    ("weld_id", 8), ("segment", 9), ("wfs", 8), ("volts", 8), ("gas_cfh", 9), ("goal", 20),
    ("method", 26), ("start_flat", 10), ("start_pos", 10), ("end_flat", 9),
    ("end_pos", 9), ("result", 18), ("noise", 7), ("noise_note", 20),
    ("notes", 50), ("before_photo", 20), ("after_photo", 20),
    ("start_time_utc", 24), ("stop_time_utc", 24), ("runtime_s", 10),
]
ROWS = 500  # pre-validated blank rows

LISTS = {
    "shift": ["day", "night"],
    "pole_side": ["top", "bottom"],
    "flange": ["top", "bottom"],
    "weld_side": ["inner", "outer"],
    "direction": ["up", "down"],
    "rotation": ["CW", "CCW"],
    "goal": ["clean", "slag inclusion", "porosity", "lack of fusion", "lack of penetration"],
    "start_pos": ["start", "mid", "end"],
    "end_pos": ["start", "mid", "end"],
    "result": ["as intended", "clean, no defect", "partial", "different defect", "not inspected"],
    "noise": ["yes", "no"],
}

wb = Workbook()

# ---------- Lists sheet (dropdown sources) ----------
ls = wb.active
ls.title = "Lists"
for c, (name, vals) in enumerate(LISTS.items(), start=1):
    ls.cell(row=1, column=c, value=name).font = Font(bold=True)
    for r, v in enumerate(vals, start=2):
        ls.cell(row=r, column=c, value=v)
    ls.column_dimensions[get_column_letter(c)].width = 20

# ---------- Weld Log sheet ----------
ws = wb.create_sheet("Weld Log", 0)
head_fill = PatternFill("solid", fgColor="1B1D21")
sess_fill = PatternFill("solid", fgColor="EEEDE9")
thin = Side(style="thin", color="D6D4CE")
idx = {}
for c, (name, width) in enumerate(COLS, start=1):
    idx[name] = c
    cell = ws.cell(row=1, column=c, value=name)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = head_fill
    cell.alignment = Alignment(vertical="center")
    ws.column_dimensions[get_column_letter(c)].width = width

# extra computed column
chk = len(COLS) + 1
ws.cell(row=1, column=chk, value="runtime_check_s").font = Font(bold=True, color="FFFFFF")
ws.cell(row=1, column=chk).fill = PatternFill("solid", fgColor="D3601A")
ws.column_dimensions[get_column_letter(chk)].width = 16
ws.cell(row=1, column=chk + 1, value="runtime_diff_s").font = Font(bold=True, color="FFFFFF")
ws.cell(row=1, column=chk + 1).fill = PatternFill("solid", fgColor="D3601A")
ws.column_dimensions[get_column_letter(chk + 1)].width = 14

sc = get_column_letter(idx["start_time_utc"])
ec = get_column_letter(idx["stop_time_utc"])
rc = get_column_letter(idx["runtime_s"])
cc = get_column_letter(chk)
last = ROWS + 1
for r in range(2, last + 1):
    # recompute runtime from the two stamps (works whether Excel kept them as
    # datetimes or text) and compare with the app's own runtime_s
    ws.cell(row=r, column=chk,
            value=f'=IF(OR({sc}{r}="",{ec}{r}=""),"",IFERROR(ROUND((VALUE({ec}{r})-VALUE({sc}{r}))*86400,3),"check"))')
    ws.cell(row=r, column=chk + 1,
            value=f'=IF(OR({cc}{r}="",{rc}{r}="",NOT(ISNUMBER({cc}{r}))),"",ROUND({cc}{r}-{rc}{r},3))')
    for name in ("session_date", "shift", "session", "unit", "pole_side", "flange",
                 "weld_side", "direction", "rotation", "setup_notes",
                 "recorder_clock_reading", "phone_time_at_reading_utc", "std_wfs", "std_volts"):
        ws.cell(row=r, column=idx[name]).fill = sess_fill

# number / date formats
for name in ("phone_time_at_reading_utc", "start_time_utc", "stop_time_utc"):
    col = get_column_letter(idx[name])
    for r in range(2, last + 1):
        ws[f"{col}{r}"].number_format = "yyyy-mm-dd hh:mm:ss.000"
for name in ("runtime_s",):
    col = get_column_letter(idx[name])
    for r in range(2, last + 1):
        ws[f"{col}{r}"].number_format = "0.000"
ws.column_dimensions[get_column_letter(idx["session_date"])].number_format = "yyyy-mm-dd"

# dropdowns
for c, (name, vals) in enumerate(LISTS.items(), start=1):
    col = get_column_letter(c)
    dv = DataValidation(type="list", formula1=f"=Lists!${col}$2:${col}${len(vals) + 1}",
                        allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    tcol = get_column_letter(idx[name])
    dv.add(f"{tcol}2:{tcol}{last}")

# table with filters, freeze panes
ref = f"A1:{get_column_letter(chk + 1)}{last}"
tab = Table(displayName="WeldLog", ref=ref)
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
ws.add_table(tab)
ws.freeze_panes = ws.cell(row=2, column=idx["weld_id"])  # session cols + header stay put
ws.row_dimensions[1].height = 22

# ---------- README ----------
rd = wb.create_sheet("README")
rd.column_dimensions["A"].width = 110
lines = [
    ("DAQ Tool workbook", True),
    ("", False),
    ("How to bring data in from the phone app", True),
    ("1. In the app, tap Copy for Excel (or open the CSV from Share bundle).", False),
    ("2. On the Weld Log sheet, click the first empty cell in column A and paste.", False),
    ("   Each value lands in its own column. The header row in the paste matches this sheet, so delete", False),
    ("   the pasted header line if you pasted it too (Copy for Excel includes one).", False),
    ("3. Dropdowns, filters, and the runtime check extend automatically to new rows inside the table.", False),
    ("", False),
    ("Columns", True),
    ("Grey columns (A to N) are session info, repeated on every weld row so filters and pivots work.", False),
    ("White columns (O onward) are per weld. One row = one weld.", False),
    ("start_time_utc / stop_time_utc / phone_time_at_reading_utc are UTC, to the millisecond, from the phone clock.", False),
    ("recorder_clock_reading is whatever the recording unit displayed when phone_time_at_reading_utc was stamped.", False),
    ("   Recorder time = phone UTC + (recorder_clock_reading - phone_time_at_reading_utc). Use that offset to line up sensor logs.", False),
    ("runtime_s is what the app measured. runtime_check_s recomputes it from the two stamps. runtime_diff_s should be 0.", False),
    ("segment is filled only when the welder stopped and restarted by accident mid-weld: each piece gets its own weld ID and a segment like 2/3.", False),
    ("before_photo / after_photo hold the file names inside the Share bundle ZIP (sN_wM_before.jpg).", False),
    ("", False),
    ("Dropdown lists live on the Lists sheet. Add a value there and it appears in the dropdown.", False),
    ("Keep this file as the template. Save each shift's data as a copy.", False),
]
for r, (text, bold) in enumerate(lines, start=1):
    c = rd.cell(row=r, column=1, value=text)
    c.font = Font(bold=bold, size=12 if bold else 11)
    c.alignment = Alignment(wrap_text=True)

wb.save("DAQTool-Template.xlsx")
print("saved DAQTool-Template.xlsx")
