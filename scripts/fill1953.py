"""Fill the 1953 columns from the United States Government Organization Manual 1953-54 (revised as of July 1, 1953;
the only edition). Run after fill1952.py. Page numbers are printed pages: PDF = printed + 9 throughout. The Post
Office chapter still describes the Department as of July 1, 1952 (p. 185 note).

Same three requirements as 1944-52: an agency name (not a position), set apart by its formatting, followed by a
description. Every row live in 1952 was matched to its 1953 heading; changed units were read on the page images.

1953 changes that matter:
  - Reorganization Plan 1 of 1953 abolished the Federal Security Agency and created the Department of Health,
    Education, and Welfare (April 11, 1953), now printed as an executive department (pp. 311-330). New row for the
    Department; the Federal Security Agency row ends; its units move under the Department.
  - Reorganization Plan 6 of 1953 (June 30) abolished the Munitions Board, the Research and Development Board, and
    the Defense Supply Management Agency and added Assistant Secretaries of Defense. The Assistant Secretaries are
    bold job titles (positions); the Office of Public Information is only "DIRECTOR, OFFICE OF PUBLIC INFORMATION.--"
    and the Office of Medical Services is replaced by an Assistant Secretary not yet described (pp. 120-125).
  - The Technical Cooperation Administration, the Advisory Committee on Voluntary Foreign Aid, and the Public Advisory
    Board for Mutual Security are printed under the Mutual Security Agency (pp. 424-428).
  - The Economic Stabilization Agency chapter now has only the National Enforcement Commission and the Office of
    Rent Stabilization (pp. 349-350); the Office of Price Stabilization and the Wage and Salary Stabilization Boards
    ended April 30, 1953.
  - State: a bold Bureau of Security and Consular Affairs (p. 77) replaces the Office of Consular Affairs, with
    Passport, Visa, Protective Services, Security, and Munitions Control offices as run-ins; the Office of Budget
    and Finance is split into Offices of Budget and Reports and of Finance (p. 77).
  - Treasury: the Analysis Staff supersedes the Tax Advisory Staff and the Office of the Technical Staff (p. 91);
    the Committee on Practice was replaced on January 9, 1953, by an Internal Revenue office (p. 98).
  - Justice: the Claims Division is now the Civil Division and absorbs the Customs Division (p. 178); a bold Office
    of Legal Counsel (p. 176) succeeds the Executive Adjudications Division, which had a bold heading in 1952 too.
Renames keep the same row: Claims Division -> Civil Division (80); Office of Foreign Agricultural Relations ->
Foreign Agricultural Service (56); Civilian Components Policy Board -> Reserve Forces Policy Board (856); Executive
Office of the Secretary (Defense) -> Office of the Secretary of Defense (983); National Capital Park and Planning
Commission -> National Capital Planning Commission (364); Office of International Trade Policy -> Office of Economic
Defense and Trade Policy (653); Office of Transportation and Communications -> Office of Transport and
Communications Policy (577); Office of Research and Intelligence -> Office of Intelligence Research (716); Office of
Intelligence Collection and Dissemination -> Office of Libraries and Intelligence Acquisition (717); Office of the
Director of Personnel (Labor) -> Office of Personnel Administration (688); Office of the Technical Staff -> Analysis
Staff (838); Office of Inter-Agency and International Relations -> Office of International Relations (802). Dead
rows reused for the new State bureau's offices: Passport Division -> Passport Office (166), Visa Division -> Visa
Office (170), Division of Protective Services -> Office of Protective Services (836), Munitions Division -> Office of
Munitions Control (763).

Corrections to earlier years (bold headings the OCR text cut short, confirmed on the page images):
  - Federal Home Loan Bank System (323): bold heading with a description under the Home Loan Bank Board in 1950-53
    (1950 p. 366, 1953 p. 408); it had been coded Dead from 1950.
  - Federal National Mortgage Association (329): bold heading under the Housing and Home Finance Agency in 1951-53
    (1951 p. 407, 1952 p. 419, 1953 p. 407); it had been coded Dead from 1951.
  - National Archives Trust Fund Board (875): "Related Organizations" heading with a description in 1951-53.
  - New rows also filled for earlier years (EARLIER): the Public Health Service's Office of the Surgeon General
    (bold heading 1950-53, same level as its Bureau of Medical Services), the Government Patents Board (own section
    1951-53), and the Office of Legal Counsel (1952 as the Executive Adjudications Division).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1952.xlsx'
DST = 'data/Spreadsheet 1935-1953.xlsx'
YEAR, PREV_YEAR = 1953, 1952

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
HEW = 'Department of Health, Education, and Welfare'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME, HEW}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
ESA, MSA, NAVY = 'Economic Stabilization Agency', 'Mutual Security Agency', 'Department of the Navy'
BSCA = 'Bureau of Security and Consular Affairs'


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

# Executive Office
rows(DEAD, 377)                                 # Liaison Office for Personnel Management: not printed
rows(DEAD, 816)                                 # National Security Resources Board: abolished, functions to ODM (p. 66)
rows(DEAD, 1002)                                # Defense Production Administration: abolished, to ODM (p. 66)
# State
rows(DEAD, 715)                                 # Office of Budget and Finance -> Offices of Budget and Reports, Finance
rows(DEAD, 861)                                 # Office of Consular Affairs -> Bureau of Security and Consular Affairs
rows(DEAD, 925)                                 # Office of German Public Affairs: not printed
rows(DEAD, 1022)                                # Management Staff: not printed
rows(BSCA, 166, 170, 836, 763)                  # Passport, Visa, Protective Services, Munitions Control run-ins (p. 78)
rows(MSA, 969, 747)                             # TCA, Advisory Committee on Voluntary Foreign Aid (p. 427)
# Treasury
rows(DEAD, 857)                                 # Tax Advisory Staff: superseded by the Analysis Staff (p. 91)
rows(DEAD, 242)                                 # Committee on Practice: replaced January 9, 1953 (p. 98)
# Defense (pp. 120-127)
rows(DEAD, 807, 808)                            # Munitions Board, Research and Development Board: abolished June 30
rows(DEAD, 854)                                 # Military Liaison Committee: named in a sentence only (p. 124)
rows(DEAD, 883)                                 # Office of Medical Services: not printed
rows(DEAD, 895)                                 # Office of Public Information: only a "DIRECTOR, ..." run-in (p. 121)
rows(DEAD, 897, 984)                            # Defense Management Council, Defense Management Staff: not printed
rows(DEAD, 986)                                 # Weapons Systems Evaluation Group: not printed
# Justice, Interior, Labor
rows(DEAD, 82)                                  # Customs Division: merged into the Civil Division (p. 178)
rows(DEAD, 957, 958)                            # Defense Electric Power, Defense Fisheries Administrations: appendix only
rows(DEAD, 687)                                 # Office of Budget and Management (Labor): not printed
# Health, Education, and Welfare: units of the Federal Security Agency move to the Department (below)
rows(DEAD, 337)                                 # Federal Security Agency: abolished April 11, 1953
# Independent agencies
rows(MSA, 1004)                                 # Public Advisory Board for Mutual Security (p. 426)
rows(DEAD, 1006, 1007, 1008)                    # Salary and Wage Stabilization Boards, Office of Price Stabilization
rows(DEAD, 822)                                 # Loyalty Review Board: not printed
rows(DEAD, 936)                                 # Contract Settlement Appeal Board: abolished January 13, 1953 (p. 392)
rows(DEAD, 652)                                 # Bureau of Community Facilities: only a division named in text (p. 406)
rows(DEAD, 851, 852)                            # Displaced Persons, Motor Carrier Claims Commissions: not printed
rows(DEAD, 1037, 1039)                          # Missouri Basin Survey Commission, National Advisory Board on
                                                # Mobilization Policy: not printed
rows(DEAD, 1013)                                # President's Advisory Committee on Management: replaced (p. 444)

# Earlier-year corrections to existing rows: {row: (years, code)}
FIX = {323: ((1950, 1951, 1952, 1953), 'Home Loan Bank Board'),
       329: ((1951, 1952, 1953), 'Housing and Home Finance Agency'),
       875: ((1951, 1952, 1953), 'National Archives and Records Service')}

EARLIER = {'Office of the Surgeon General (Public Health Service)': (1950, 1951, 1952),
           'Government Patents Board': (1951, 1952), 'Office of Legal Counsel': (1952,)}

NEW = [
    (HEW, CAB),                                                    # p. 311
    ('Office of the Surgeon General (Public Health Service)', 'Public Health Service'),   # p. 314
    # State (p. 77)
    (BSCA, ST), ('Office of Security (State)', BSCA),
    ('Office of Budget and Reports (State)', ST), ('Office of Finance (State)', ST),
    # Justice, Interior, Commerce, Labor
    ('Office of Legal Counsel', JUS),                              # p. 176
    ('Technical Review Staff (Interior)', INT),                    # p. 200
    ('Office of General Counsel (Commerce)', COM),                 # p. 262
    ('Office of the Administrative Assistant Secretary (Labor)', LAB),   # p. 290
    # Defense
    ('Office of Information (Navy)', NAVY),                        # p. 157
    # Independent agencies
    ('National Enforcement Commission', ESA),                      # p. 349
    ('Federal Coal Mine Safety Board of Review', OTHER),           # p. 354
    ('Government Patents Board', OTHER),                           # p. 401
    ('Interdepartmental Committee on Scientific Research and Development', OTHER),   # p. 415
    ("President's Advisory Committee on Government Organization", OTHER),            # p. 444
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
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}

last = max(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value)
for r in range(3, last + 1):
    if not ws.cell(r, 1).value:
        continue
    prev = [ws.cell(r, PREV + i).value for i in range(5)]
    if r in O:
        new = vals(O[r])
    elif all(v in (None, '') for v in prev):
        continue
    elif prev[3] == 'Federal Security Agency':
        new = vals(HEW)
    else:
        new = prev
    for i, v in enumerate(new):
        ws.cell(r, COL + i).value = v

for r, (years, c) in FIX.items():
    for y in years:
        for i, v in enumerate(vals(c)):
            ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v

r = last + 1
for name, c in NEW:
    ws.cell(r, 1).value = name
    ws.cell(r, 2).value = 'Executive'
    for col in range(3, COL):
        ws.cell(r, col).value = vals(NOTYET)[(col - 3) % 5]
    for y in EARLIER.get(name, ()) + (YEAR,):
        for i, v in enumerate(vals(c)):
            ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v
    r += 1

# FAQ: hide (never delete) dead rows and legislative/judicial rows
for x in range(3, ws.max_row + 1):
    if ws.cell(x, 1).value:
        kind = (ws.cell(x, 2).value or '').strip()
        ws.row_dimensions[x].hidden = kind in ('Legislative', 'Judicial') or str(ws.cell(x, COL).value).strip() == 'Dead'

wb.save(DST)
print('saved', DST, 'new rows', last + 1, '-', r - 1)
