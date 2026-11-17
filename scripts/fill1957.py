"""Fill the 1957 columns from the United States Government Organization Manual 1957-58 (revised as of June 1, 1957;
the only edition). Run after fill1956.py and remove_org_units_1944_1956.py. Page numbers are printed pages:
PDF = printed + 10 through p. 108, + 8 for pp. 111-513, and + 10 from p. 514 (the scan skips pp. 109-110, which are
Internal Revenue district lists, and repeats pp. 512-513).

Same three requirements as 1944-56: an agency name (not a position), set apart by its formatting, followed by a
description. No organizational units (user decision, 2026-11-10). Every row live in 1956 was matched to its 1957
heading; changed units were read on the page images.

1957 changes that matter:
  - New "Guide to Selected Boards, Committees, and Commissions" (pp. 521-529): capitals run-ins with descriptions for
    boards no longer described in the agency chapters. User decision (2026-11-17): rows already in the sheet that
    moved there stay live with their coding (Air Coordinating Committee, Migratory Bird Conservation Commission,
    National Park Trust Fund Board, Board on Geographic Names, Government Patents Board, and others); guide entries
    not already in the sheet are not added.
  - Defense: new bold National Security Agency, "an element of the Department of Defense" (p. 137).
  - Housing and Home Finance Agency: new Federal Flood Indemnity Administration (September 1956, p. 417).
  - Interior: the United States Fish and Wildlife Service (Fish and Wildlife Act of 1956, p. 225); its Bureaus of
    Commercial Fisheries and of Sport Fisheries and Wildlife are only named in a sentence, so they are not added.
  - Not printed: National Security Training Commission (Appendix A only, p. 648); Rubber Producing Facilities
    Disposal Commission (Federal Facilities Corporation is its successor, p. 102).
  - Farm Credit: the production credit corporations merged into the intermediate credit banks January 1, 1957; the
    PRODUCTION CREDIT ASSOCIATIONS heading continues on row 428 (p. 359).
Renames keep the same row: Fish and Wildlife Service -> United States Fish and Wildlife Service (190); Emergency
Procurement Service -> Defense Materials Service (1011, September 7, 1956, p. 405).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1956.xlsx'
DST = 'data/Spreadsheet 1935-1957.xlsx'
YEAR, PREV_YEAR = 1957, 1956

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
HEW = 'Department of Health, Education, and Welfare'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME, HEW}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
HHFA = 'Housing and Home Finance Agency'


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

# Independent agencies
rows(DEAD, 1040)                                # National Security Training Commission: Appendix A only (p. 648)
rows(DEAD, 1084)                                # Rubber Producing Facilities Disposal Commission: not printed

NEW = [
    ('National Security Agency', NME),                                                 # p. 137
    ('Federal Flood Indemnity Administration', HHFA),                                  # p. 417
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
