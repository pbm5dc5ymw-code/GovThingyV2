"""Fill the 1954 columns from the United States Government Organization Manual 1954-55 (revised as of July 1, 1954;
the only edition). Run after fill1953.py. Page numbers are printed pages: PDF = printed + 8 throughout.

Same three requirements as 1944-53: an agency name (not a position), set apart by its formatting, followed by a
description. Every row live in 1953 was matched to its 1954 heading; changed units were read on the page images.

1954 changes that matter:
  - Reorganization Plans 7 and 8 of 1953 (August 1): the Foreign Operations Administration replaces the Mutual
    Security Agency, the Technical Cooperation Administration, and the Institute of Inter-American Affairs; its
    Public Advisory Board, International Development Advisory Board, Advisory Committee on Voluntary Foreign Aid,
    Labor Advisory Committee, and Health Advisory Committee have bold headings (pp. 385-389). The United States
    Information Agency takes over the International Information Administration; its media services are described
    only in a plain paragraph (pp. 485-488), and the educational exchange program stays in State as plain text.
  - New independent agencies: Small Business Administration (replaces the Small Defense Plants Administration and
    the RFC's lending), Advisory Committee on Weather Control, Rubber Producing Facilities Disposal Commission,
    Saint Lawrence Seaway Development Corporation, Government Contract Committee, Foreign Claims Settlement
    Commission. Ended: Economic Stabilization Agency and its units, Defense Materials Procurement Agency (to GSA),
    Reconstruction Finance Corporation (abolished June 30, 1954), Council of National Defense (appendix only),
    Office for Emergency Management, Federal Personnel Council, Office of Contract Settlement, Home Owners' Loan
    Corporation ("since liquidated", p. 408). The Farm Credit Administration is independent again (Farm Credit Act
    of 1953, p. 341); its Cooperative Research and Service Division moved to Agriculture as the Farmer Cooperative
    Service (p. 233).
  - Agriculture (Secretary's Memorandum 1320, Supplement 4, November 2, 1953, pp. 223-250): the Agricultural Research
    Service replaces the Agricultural Research Administration and its bureaus (the Office of Experiment Stations and
    the Agricultural Research Center keep small bold headings inside it); the Agricultural Marketing Service and the
    Commodity Stabilization Service replace the Production and Marketing Administration and the Bureau of
    Agricultural Economics; new Agricultural Conservation Program Service.
  - Commerce: the Business and Defense Services Administration replaces the National Production Authority; new
    Bureau of Foreign Commerce, Office of the Under Secretary, and Appeals Board; the Office of Business Economics is
    a bold section again (p. 258). The Inland Waterways Corporation is no longer described.
  - State: the Bureau of German Affairs is now an Office of German Affairs in the Bureau of European Affairs; the UN
    bureau's Office of International Administration and Conferences is split in two; new run-in Division of
    Cryptography and Office of Communications and Records; the Office of Public Affairs, UNESCO Relations Staff,
    and Foreign Service Inspection Corps are no longer set off (p. 81 has only "INSPECTION.--").
  - Post Office (pp. 193-198): new Bureau of Personnel, Office of the Special Assistant to the Postmaster General,
    and Office of the Executive Assistant to the Deputy Postmaster General; the Purchasing Agent, Chief Clerk and
    Director of Personnel, and Administrative Assistant offices are gone.
  - HEW (pp. 303-304): Office of Internal Security and Departmental Council are new run-ins; the Offices of Field
    Services, International Relations, and Federal-State Relations are gone.
  - Executive Office: new Operations Coordinating Board, printed with the Central Intelligence Agency under the
    National Security Council (p. 67).
Renames keep the same row: Bureau of Internal Revenue -> Internal Revenue Service (239); Bureau of Security and
Consular Affairs -> Bureau of Inspection, Security and Consular Affairs (1045); Office of Protective Services ->
Office of Special Consular Services (836); Office of Operating Facilities -> Office of General Services (862); Bureau
of Accounts (Post Office) -> Bureau of the Controller (424); Extension Service -> Federal Extension Service (626);
Cooperative Research and Service Division -> Farmer Cooperative Service (302); Columbia Institution for the Deaf ->
Gallaudet College (849). Dead rows reused: Agricultural Marketing Service (29), Office of Business Economics (703).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1953.xlsx'
DST = 'data/Spreadsheet 1935-1954.xlsx'
YEAR, PREV_YEAR = 1954, 1953

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
HEW = 'Department of Health, Education, and Welfare'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME, HEW}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
FOA, ARS = 'Foreign Operations Administration', 'Agricultural Research Service'


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
rows(DEAD, 381)                                 # Office for Emergency Management: not printed
# State (pp. 79-87)
rows(DEAD, 655, 879)                            # Office of Public Affairs, UNESCO Relations Staff: plain text (p. 85)
rows(DEAD, 827)                                 # International Educational Exchange Service: plain paragraph (p. 85)
rows(DEAD, 1021)                                # Foreign Service Inspection Corps: only "INSPECTION.--" (p. 81)
rows(DEAD, 904, 923, 924)                       # Bureau of German Affairs and its offices -> Office of German Affairs
rows(DEAD, 921)                                 # Office of International Administration and Conferences: split
rows(DEAD, 1023, 1024, 1025, 1026, 1027)        # International Information Administration and services -> USIA
# Post Office (pp. 193-198)
rows(DEAD, 421, 640, 980)                       # Purchasing Agent, Chief Clerk and Director of Personnel, Admin. Asst.
# Interior
rows(DEAD, 961)                                 # Petroleum Administration for Defense: not printed
# Agriculture
rows(DEAD, 35, 34, 36, 39, 40, 41, 42)          # Agricultural Research Administration and its bureaus -> ARS
rows(ARS, 51, 631)                              # Office of Experiment Stations, Agricultural Research Center (p. 232)
rows(DEAD, 32, 651)                             # Bureau of Agricultural Economics, Production and Marketing Admin.
rows(AG, 29)                                    # Agricultural Marketing Service (p. 238)
rows(AG, 302)                                   # Farmer Cooperative Service (p. 233)
rows(OTHER, 305)                                # Farm Credit Administration: independent again (p. 341)
# Commerce
rows(DEAD, 962)                                 # National Production Authority -> BDSA (p. 257)
rows(DEAD, 352)                                 # Inland Waterways Corporation: named in a list only (p. 253)
rows(COM, 703)                                  # Office of Business Economics (p. 258)
# Labor, HEW
rows(DEAD, 982)                                 # Defense Manpower Administration: not printed
rows(DEAD, 933, 802, 801)                       # Offices of Field Services, International Relations, Federal-State
# Independent agencies
rows(DEAD, 1038, 969, 878, 1028, 1029)          # Mutual Security Agency, TCA and its units -> FOA
rows(FOA, 1004, 747)                            # Public Advisory Board, Advisory Committee on Voluntary Foreign Aid
rows(DEAD, 1005, 1036, 1054)                    # Economic Stabilization Agency and its units
rows(DEAD, 1035)                                # Defense Materials Procurement Agency: functions to GSA
rows(DEAD, 1041)                                # Small Defense Plants Administration -> Small Business Administration
rows(DEAD, 432)                                 # Reconstruction Finance Corporation: not printed
rows(DEAD, 462)                                 # Council of National Defense: appendix only
rows(DEAD, 26, 519)                             # Federal Personnel Council (both rows): not printed
rows(DEAD, 565)                                 # Office of Contract Settlement: named in GSA history only (p. 392)
rows(DEAD, 349)                                 # Home Owners' Loan Corporation: "since liquidated" (p. 408)

NEW = [
    ('Operations Coordinating Board', 'National Security Council'),                  # p. 67
    # State
    ('Division of Cryptography (State)', ST), ('Office of Communications and Records (State)', ST),   # pp. 80-81
    ('Office of German Affairs', 'Bureau of European Affairs'),                      # p. 83
    ('Office of International Conferences', 'Bureau of United Nations Affairs'),     # p. 84
    ('Office of International Administration', 'Bureau of United Nations Affairs'),  # p. 84
    # Post Office (pp. 194, 197)
    ('Office of the Special Assistant to the Postmaster General', PO),
    ('Office of the Executive Assistant to the Deputy Postmaster General', PO),
    ('Bureau of Personnel (Post Office)', PO),
    # Agriculture (pp. 224-240)
    ('Agricultural Conservation Program Service', AG), (ARS, AG), ('Commodity Stabilization Service', AG),
    # Commerce (pp. 254-266)
    ('Office of the Under Secretary of Commerce', COM), ('Appeals Board (Commerce)', COM),
    ('Business and Defense Services Administration', COM), ('Bureau of Foreign Commerce', COM),
    # Labor, HEW
    ("Office of the President's Committee on Employment of the Physically Handicapped", LAB),   # p. 283
    ('Office of Internal Security (HEW)', HEW), ('Departmental Council (HEW)', HEW),            # pp. 303-304
    # Independent agencies
    ('Advisory Committee on Weather Control', OTHER),                                # p. 323
    (FOA, OTHER),                                                                    # p. 385
    ('International Development Advisory Board', FOA),                               # p. 388
    ('Labor Advisory Committee (FOA)', FOA), ('Health Advisory Committee (FOA)', FOA),   # p. 389
    ('Government Contract Committee', OTHER),                                        # p. 401
    ('Rubber Producing Facilities Disposal Commission', OTHER),                      # p. 449
    ('Saint Lawrence Seaway Development Corporation', OTHER),                        # p. 450
    ('Small Business Administration', OTHER),                                        # p. 461
    ('United States Information Agency', OTHER),                                     # p. 485
    ('Foreign Claims Settlement Commission of the United States', OTHER),            # p. 513
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
