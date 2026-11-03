"""Fill the 1952 columns from the United States Government Organization Manual 1952-53 (revised as of July 1, 1952;
the only edition). Run after fill1951.py. Page numbers are printed pages: PDF = printed + 12 for pp. 57-385 and
+ 10 from p. 388; the scan has no pages 386-387 (inside the Federal Security Agency's Social Security section).

Same three requirements as 1944-51: an agency name (not a position), set apart by its formatting, followed by a
description. Every row live in 1951 was matched to its 1952 heading; changed units were read on the page images.

1952 layout changes that matter:
  - Commerce now heads its staff with bold "Office of Assistant Secretary of Commerce for ..." sections and describes
    the offices under them (and the Bureau of Foreign and Domestic Commerce) only in plain paragraphs (pp. 260-264).
  - State: the International Information Administration (January 1952) replaces the Offices of International
    Information and Educational Exchange; the Technical Cooperation Administration now has development services and
    the Institute of Inter-American Affairs under it (pp. 80-83).
  - The Mutual Security Agency replaces the Economic Cooperation Administration (abolished by act of October 10, 1951).
  - The Office of Rent Stabilization (Economic Stabilization Agency) takes over the Office of the Housing Expediter.
Renames keep the same row: Defense Minerals -> Defense Minerals Exploration Administration (959); Defense Management
Committee -> Council (897); Office of Management Engineering -> Defense Management Staff (984); Air Proving Ground ->
Air Proving Ground Command (990); Office of Educational Exchange -> International Educational Exchange Service (827);
Public Advisory Board (ECA) -> Public Advisory Board, Mutual Security Program (1004, now in the Executive Office).

Corrections to earlier years: the 1951 Navy chapter already has bold "Office of Naval Material" and "Office of Naval
Research" headings (1951 pp. 148-149), and the 1950 Army chapter has a bold "Army Staff" heading (1950 p. 126; not
in 1951); those new rows are filled for those years too (EARLIER).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1951.xlsx'
DST = 'data/Spreadsheet 1935-1952.xlsx'
YEAR, PREV_YEAR = 1952, 1951

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME}
DEAD, NOTYET, OTHER = 'DEAD', 'NOTYET', 'OTHER'
EOP, ESA = 'Executive Office of the President', 'Economic Stabilization Agency'
ARMY, NAVY = 'Department of the Army', 'Department of the Navy'
TCA, IIA = 'Technical Cooperation Administration', 'International Information Administration'


def vals(c):
    """Five-column coding: cabinet, child, cabinet parent, other parent, other agency."""
    if c == DEAD: return ['Dead'] * 5
    if c == NOTYET: return ['No', 'No', 'None', 'None', 'No']
    if c == OTHER: return ['No', 'No', 'None', 'None', 'Yes']
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


O = {}
def rows(code, *row_numbers):
    O.update({r: code for r in row_numbers})

# Commerce: offices described only in plain paragraphs under the Assistant Secretaries' sections (pp. 260-264)
rows(DEAD, 63, 703, 704, 705, 706)              # Bureau of Foreign and Domestic Commerce and its offices
rows(DEAD, 790, 931, 932, 965, 966, 967, 968)   # Technical Services, Program Planning, Publications, Budget and
                                                # Management, Personnel, Administrative Services, Transportation
rows(DEAD, 963)                                 # Industry Evaluation Board: members list and a sentence (pp. 259, 261)
# State
rows(DEAD, 826)                                 # Office of International Information -> IIA services (p. 83)
rows(DEAD, 971)                                 # Program Planning and Evaluation Staff: not printed
rows(TCA, 878)                                  # Institute of Inter-American Affairs: run-in under TCA (p. 81)
rows(IIA, 827)                                  # International Educational Exchange Service: run-in under IIA (p. 83)
# Defense
rows(DEAD, 985)                                 # Military Traffic Service: now only "DIRECTOR, MILITARY TRAFFIC SERVICE.--" (p. 123)
rows(DEAD, 997)                                 # Special Weapons Command: not printed
# Executive Office and independent agencies
rows(EOP, 1004)                                 # Public Advisory Board, Mutual Security Program: heading in the Executive Office (p. 72)
rows(DEAD, 817)                                 # Economic Cooperation Administration: abolished October 10, 1951
rows(DEAD, 748)                                 # Office of the Housing Expediter: appendix only; -> Office of Rent Stabilization
rows(DEAD, 1010)                                # GSA Office of Defense Coordination: not printed

EARLIER = {'Office of Naval Material': (1951,), 'Office of Naval Research': (1951,), 'Army Staff': (1950,)}

NEW = [
    # Commerce (pp. 260-263, 277)
    ('Office of the Under Secretary of Commerce for Transportation', COM),
    ('Office of the Assistant Secretary of Commerce for Domestic Affairs', COM),
    ('Office of the Assistant Secretary of Commerce for International Affairs', COM),
    ('Office of the Assistant Secretary of Commerce for Administration', COM),
    ('National Shipping Authority', 'Maritime Administration'),
    # State (pp. 79-83)
    ('Foreign Service Inspection Corps', ST), ('Management Staff (State)', ST),
    (IIA, ST), ('International Broadcasting Service', IIA), ('International Press Service', IIA),
    ('International Motion Picture Service', IIA), ('International Information Center Service', IIA),
    ('Near East and African Development Service', TCA), ('Asian Development Service', TCA),
    # Treasury (p. 93)
    ('Office of Administrative Assistant Secretary (Treasury)', TRE),
    # Defense
    ('Army Staff', ARMY),                                          # p. 135
    ('Office of Naval Material', NAVY), ('Office of Naval Research', NAVY),   # p. 156
    ('Office of the Comptroller (Navy)', NAVY),                    # p. 157
    # Independent agencies
    ('Defense Materials Procurement Agency', OTHER),               # p. 323
    ('Office of Rent Stabilization', ESA),                         # p. 342
    ('Missouri Basin Survey Commission', OTHER),                   # p. 435
    ('Mutual Security Agency', OTHER),                             # p. 437
    ('National Advisory Board on Mobilization Policy', OTHER),     # p. 439
    ('National Security Training Commission', OTHER),              # p. 453
    ('Small Defense Plants Administration', OTHER),                # p. 476
    ("Veterans' Education Appeals Board", OTHER),                  # p. 519
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
    for y in EARLIER.get(name, ()) + (YEAR,):
        for i, v in enumerate(vals(c)):
            ws.cell(r, hdr.get(f'{y}_cabinet_agency', COL) + i).value = v
    r += 1

# FAQ: hide (never delete) dead rows and legislative/judicial rows
for x in range(3, ws.max_row + 1):
    if ws.cell(x, 1).value:
        kind = (ws.cell(x, 2).value or '').strip()
        ws.row_dimensions[x].hidden = kind in ('Legislative', 'Judicial') or str(ws.cell(x, COL).value).strip() == 'Dead'

wb.save(DST)
print('saved', DST, 'new rows', last + 1, '-', r - 1)
