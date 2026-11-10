"""Apply the "actual agencies, not organizational units" rule (user decision, 2026-11-10) to 1944-1956.

Run after fill1956.py, on the 1935-1956 workbook only. Decisions come from data/Organizational units review
1944-1956.xlsx: a row is removed when its "Your decision" is Remove, or when that cell is blank and the suggestion is
Remove or Check (the user chose to remove every Check row). Removed rows are set to Dead in every year from 1944 to
1956 in which they were live, and hidden if Dead in the latest year; nothing is deleted. 1935-1943 (the user's
hand-keyed years) are not touched.
"""
import openpyxl

SHEET = 'data/Spreadsheet 1935-1956.xlsx'
REVIEW = 'data/Organizational units review 1944-1956.xlsx'
YEARS = range(1944, 1957)

review = openpyxl.load_workbook(REVIEW).active
REMOVE = set()
for _, row, _, _, _, suggestion, _, decision in review.iter_rows(min_row=3, values_only=True):
    if decision == 'Remove' or (not decision and suggestion in ('Remove', 'Check')):
        REMOVE.add(row)

wb = openpyxl.load_workbook(SHEET)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
changed = 0
for r in sorted(REMOVE):
    for y in YEARS:
        C = hdr[f'{y}_cabinet_agency']
        v = [str(ws.cell(r, C + i).value or '').strip() for i in range(5)]
        if 'Yes' in (v[0], v[1], v[4]):
            for i in range(5):
                ws.cell(r, C + i).value = 'Dead'
            changed += 1

# FAQ: hide (never delete) dead rows and legislative/judicial rows
LAST = hdr[f'{max(YEARS)}_cabinet_agency']
for x in range(3, ws.max_row + 1):
    if ws.cell(x, 1).value:
        kind = (ws.cell(x, 2).value or '').strip()
        ws.row_dimensions[x].hidden = kind in ('Legislative', 'Judicial') or str(ws.cell(x, LAST).value).strip() == 'Dead'

wb.save(SHEET)
print(f'{len(REMOVE)} rows, {changed} row-years set to Dead')
