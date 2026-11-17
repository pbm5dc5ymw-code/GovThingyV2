"""Fill the 1958 columns from the United States Government Organization Manual 1958-59 (revised as of June 1, 1958;
the only edition). Run after fill1957.py. Page numbers are printed pages: PDF = printed + 11 throughout the parts
used (the offset changes only in the international section, p. 566).

Same three requirements as 1944-57: an agency name (not a position), set apart by its formatting, followed by a
description. No organizational units (user decision, 2026-11-10); Guide to Selected Boards, Committees, and
Commissions: rows already in the sheet stay, no new rows from it (user decision, 2026-11-17). Every row live in 1957
was matched to its 1958 heading; changed units were read on the page images.

1958 changes that matter:
  - Treasury: the Office of Production and Defense Lending is replaced by the Office of Defense Lending (July 1, 1957,
    Treasury Order 185, p. 103), which administers the Reconstruction Finance Corporation Liquidation Fund; the RFC
    itself and the Federal Facilities Corporation are no longer printed.
  - State: new bold Bureau of Public Affairs (p. 84), a bureau like the geographic bureaus. Not added: the Bureau of
    Administration (the department's administrative area, p. 89), the Development Loan Fund (a program run-in in the
    ICA section, p. 93), and the temporary Office of the U. S. Commissioner General for the Brussels Exhibition (p. 94).
  - Defense: new Advanced Research Projects Agency (February 7, 1958, p. 193).
  - Post Office: the Bureau of Post Office Operations is now the Bureau of Operations, the sole operating bureau
    (p. 213).
  - National Archives and Records Service: "PRESIDENTIAL LIBRARIES (FRANKLIN D. ROOSEVELT LIBRARY; HARRY S. TRUMAN
    LIBRARY).—" run-in (p. 420), so the Truman Library is added beside the Roosevelt Library.
  - Independent agencies: new Airways Modernization Board (Act of August 14, 1957, p. 352). The Federal Flood
    Indemnity Administration is not printed. The Commission on Civil Rights (1957) appears only in the Guide (p. 533),
    so it is not added under the Guide rule. Justice's new Civil Rights Division is a litigating division (not added).
  - Verification pass (user request, 2026-11-17): every live row was checked on the page image, with the printed
    page recorded in data/Evidence 1957-1958.xlsx. The Bureau of Economic Affairs has its own bold heading again
    (p. 88) and comes back on its row; the Veterans' Education Appeals Board and the Advisory Committee on Weather
    Control are no longer printed. The Brussels Commissioner General's office (added in 1957) is still a bold
    heading with a description (p. 94).
Renames keep the same row: Office of Production and Defense Lending -> Office of Defense Lending (1091); Bureau of
Post Office Operations -> Bureau of Operations (941).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1957.xlsx'
DST = 'data/Spreadsheet 1935-1958.xlsx'
YEAR, PREV_YEAR = 1958, 1957

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
HEW = 'Department of Health, Education, and Welfare'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME, HEW}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
NARS = 'National Archives and Records Service'


def vals(c):
    """Five-column coding: cabinet, child, cabinet parent, other parent, other agency."""
    if c == DEAD: return ['Dead'] * 5
    if c == NOTYET: return ['No', 'No', 'None', 'None', 'No']
    if c == OTHER: return ['No', 'No', 'None', 'None', 'Yes']
    if c == CAB: return ['Yes', 'No', 'None', 'None', 'No']
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


O = {}
def rows(code, *row_numbers):
    O.update({r: code for r in row_numbers})

# Treasury
rows(DEAD, 432, 1092)                           # Reconstruction Finance Corporation, Federal Facilities Corporation:
                                                # not printed; RFC Liquidation Fund under Defense Lending (p. 103)
# State
rows(ST, 1114)                                  # Bureau of Economic Affairs: bold heading again (p. 88)
# Independent agencies
rows(DEAD, 1121)                                # Federal Flood Indemnity Administration: not printed
rows(DEAD, 1042, 1078)                          # Veterans' Education Appeals Board, Advisory Committee on Weather
                                                # Control: no longer in the Guide or elsewhere

NEW = [
    ('Bureau of Public Affairs (State)', ST),                                          # p. 84
    ('Advanced Research Projects Agency', NME),                                        # p. 193
    ('Harry S. Truman Library', NARS),                                                 # p. 420
    ('Airways Modernization Board', OTHER),                                            # p. 352
]

shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
PREV = hdr[f'{PREV_YEAR}_cabinet_agency']
COL = PREV + 5
for i, s in enumerate(['cabinet_agency', 'child_agency', 'cabinet_parent_agency', 'other_parent_agency', 'other_agency']):
    ws.cell(2, COL + i).value = f'{YEAR}_{s}'
    ws.column_dimensions[openpyxl.utils.get_column_letter(COL + i)].width = \
        ws.column_dimensions[openpyxl.utils.get_column_letter(PREV + i)].width

last = max(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value)
for r in range(3, last + 1):
    if not ws.cell(r, 1).value:
        continue
    prev = [ws.cell(r, PREV + i).value for i in range(5)]
    if r in O:
        new = vals(O[r])
    elif all(v in (None, '') for v in prev):
        continue
    else:
        new = prev
    for i, v in enumerate(new):
        ws.cell(r, COL + i).value = v

r = last + 1
for name, c in NEW:
    ws.cell(r, 1).value = name
    ws.cell(r, 2).value = 'Executive'
    for col in range(3, COL):
        ws.cell(r, col).value = vals(NOTYET)[(col - 3) % 5]
    for i, v in enumerate(vals(c)):
        ws.cell(r, COL + i).value = v
    r += 1

# FAQ: hide (never delete) dead rows and legislative/judicial rows
for x in range(3, ws.max_row + 1):
    if ws.cell(x, 1).value:
        kind = (ws.cell(x, 2).value or '').strip()
        ws.row_dimensions[x].hidden = kind in ('Legislative', 'Judicial') or str(ws.cell(x, COL).value).strip() == 'Dead'

wb.save(DST)
print('saved', DST, 'new rows', last + 1, '-', r - 1)
