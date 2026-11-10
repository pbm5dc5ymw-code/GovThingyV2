"""Fill the 1956 columns from the United States Government Organization Manual 1956-57 (revised as of June 1, 1956;
the only edition). Run after fill1955.py. Page numbers are printed pages: PDF = printed + 8 through p. 179 and
+ 10 from p. 180 (the scan repeats pp. 178-179).

Same three requirements as 1944-55: an agency name (not a position), set apart by its formatting, followed by a
description. Every row live in 1955 was matched to its 1956 heading; changed units were read on the page images.

1956 changes that matter:
  - The Foreign Operations Administration was abolished June 30, 1955; the International Cooperation Administration
    is a bold section inside State (pp. 89-90). The FOA's committees are gone; the International Development Advisory
    Board and the Advisory Committee on Voluntary Foreign Aid appear only as members lists (p. 77).
  - State: new bold Bureau of Economic Affairs (p. 86) over the Offices of International Financial and Development
    Affairs, International Trade and Resources, and Transport and Communications. The Policy Planning Staff and the
    Executive Secretariat are now described only under an Assistant Secretary heading and a "DIRECTOR, ..." run-in
    (pp. 80, 83). New offices and staffs inside bureaus or the Controller's area are organizational units, not
    agencies, so they are not added (user decision, 2026-11-10): Special Projects Staff (p. 81), Office of
    Inter-American Regional Economic Affairs (p. 84), Office of Southwest Pacific Affairs (p. 85), Personnel
    Projects Staff, Office of Deputy Assistant Secretary for Operations, Clemency and Parole Board (pp. 86-87),
    Office of Refugee and Migration Affairs (p. 89). Units carried from earlier years keep their coding.
  - Post Office: "BUREAU OF THE CHIEF POSTAL INSPECTOR.--" is a run-in with a description again (p. 204).
  - Navy: new bold Naval Research Advisory Committee (p. 168). HEW: the Offices of Inspection and of the Comptroller
    are not printed (p. 320).
  - Independent agencies: the Federal Home Loan Bank Board is independent (August 11, 1955, p. 381). The Civil
    Defense Coordinating Board and a new National Civil Defense Advisory Council are printed under the Federal Civil
    Defense Administration (p. 371). New: GSA's Transportation and Public Utilities Service (July 1, 1955, p. 422),
    and the United States Advisory Commissions on Educational Exchange and on Information (pp. 504-505).
Renames keep the same row: Home Loan Bank Board -> Federal Home Loan Bank Board (820); Bureau of the Chief Inspector
-> Bureau of the Chief Postal Inspector (422); Saline Water Conversion Program Staff -> Office of Saline Water (1097);
Office of Regional American Affairs -> Office of Inter-American Regional Political Affairs (912); Office of
Philippine and Southeast Asian Affairs -> Office of Southeast Asian Affairs (915).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1955.xlsx'
DST = 'data/Spreadsheet 1935-1956.xlsx'
YEAR, PREV_YEAR = 1956, 1955

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
HEW = 'Department of Health, Education, and Welfare'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME, HEW}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
BEA, FCDA, NAVY = 'Bureau of Economic Affairs', 'Federal Civil Defense Administration', 'Department of the Navy'
GSA = 'General Services Administration'


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

# State
rows(DEAD, 754, 823)                            # Executive Secretariat, Policy Planning Staff: under a "DIRECTOR, ..."
                                                # run-in and an Assistant Secretary heading only (pp. 80, 83)
rows(BEA, 654, 653, 577)                        # offices of the new Bureau of Economic Affairs (p. 86)
# Post Office
rows(PO, 422)                                   # Bureau of the Chief Postal Inspector run-in (p. 204)
# HEW
rows(DEAD, 1105, 1106)                          # Offices of Inspection and of the Comptroller: not printed (p. 320)
# Foreign Operations Administration (abolished June 30, 1955) and its units
rows(DEAD, 1079, 1080, 1081, 1082, 1108, 1109, 747)
# Independent agencies
rows(OTHER, 820)                                # Federal Home Loan Bank Board: independent (p. 381)
rows(FCDA, 1107)                                # Civil Defense Coordinating Board: in the FCDA section (p. 371)

NEW = [
    # State (pp. 81-89)
    ('International Cooperation Administration', ST), (BEA, ST),
    # Navy (p. 168)
    ('Naval Research Advisory Committee', NAVY),
    # Independent agencies
    ('National Civil Defense Advisory Council', FCDA),                                  # p. 371
    ('Transportation and Public Utilities Service', GSA),                               # p. 422
    ('United States Advisory Commission on Educational Exchange', OTHER),               # p. 504
    ('United States Advisory Commission on Information', OTHER),                        # p. 505
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
