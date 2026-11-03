"""Fill the 1951 columns from the United States Government Organization Manual 1951-52 (revised as of July 1, 1951;
the only edition). Run after agency_not_position_1944_1950.py. Page numbers are printed pages (PDF = printed + 8).

Rule (user, 2026-11-03): a row-year needs (1) an agency name, not a position, (2) set apart by its formatting
(bold heading, or a capitals-and-small-capitals run-in), (3) followed by a description of the agency.
Every row live in 1950 was matched to its 1951 heading; units that changed were read on the page images.

1951 layout changes that matter:
  - Interior and Commerce now describe the Secretary's staff under bold job titles ("Solicitor", "Director of
    Information"); the old staff divisions appear only in sentences, so those rows end (Interior p. 189-192).
  - Korean-war mobilization agencies: Office of Defense Mobilization (EOP), Defense Production Administration,
    Economic Stabilization Agency, Federal Civil Defense Administration, Defense Transport Administration, Interior's
    five defense administrations, Commerce's National Production Authority, Labor's Defense Manpower Administration.
  - State has a "Committees and Commissions" section (p. 84-86).
Renames keep the same row (codebook rule): Virgin Islands Company -> Corporation (858); Beltsville Research Center ->
Agricultural Research Center (631); Office of Field Operations -> Office of Field Service (704); Office of
International Trade Policy -> Office of Economic Defense and Trade Policy (653); Office of Consular Affairs -> Office
of Security and Consular Affairs (861); Labor Chief Clerk and Budget Officer, Director of Personnel, Director of
Information -> Offices of Budget and Management, Personnel Administration, Information (687-689); Civilian
Components Policy Board -> Reserve Forces Policy Board (856); Office of Medical Services -> Armed Forces Medical
Policy Council (883); The Panama Canal -> Canal Zone Government (457); Panama Railroad Company -> Panama Canal
Company (893, now its own agency); National Archives Council -> Federal Records Council (361, a GSA related body);
Office of Inter-Agency and International Relations -> Office of International Relations (802).

Some new rows were printed the same way in 1950 and never got rows; they are filled for 1950 too (ALSO_1950).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1950.xlsx'
DST = 'data/Spreadsheet 1935-1951.xlsx'
YEAR, PREV_YEAR = 1951, 1950

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME}
DEAD, NOTYET, OTHER = 'DEAD', 'NOTYET', 'OTHER'
EOP, AF, GSA = 'Executive Office of the President', 'Department of the Air Force', 'General Services Administration'
NARS, ESA, ECA = 'National Archives and Records Service', 'Economic Stabilization Agency', 'Economic Cooperation Administration'
ARAB, NEAB = 'Bureau of Inter-American Affairs', 'Bureau of Near Eastern, South Asian, and African Affairs'


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

# Interior: staff divisions now only sentences under bold position headings (pp. 189-192)
rows(DEAD, 183, 186, 198, 199, 783, 784, 840)   # Power, Information, Land Utilization, Solicitor, Oil and Gas, Geography, Programs
rows(DEAD, 928)                                 # bold "Administrative Assistant Secretary": a position (p. 191)
rows(DEAD, 370)                                 # National Power Policy Committee: not printed
# Commerce
rows(DEAD, 452)                                 # Textile Foundation: not printed
# State
rows(DEAD, 898)                                 # Mutual Defense Assistance Program: no longer a State unit; sentences only
rows(DEAD, 909, 910)                            # East Coast, North and West Coast Affairs -> Office of South American Affairs (p. 81)
rows(DEAD, 917)                                 # African and Near Eastern Affairs -> separate Near Eastern and African offices (p. 82)
rows(ST, 747)                                   # Advisory Committee on Voluntary Foreign Aid: heading in Committees and Commissions (p. 84); FAQ: unhide
# Treasury
rows(DEAD, 263)                                 # Tax Legislative Counsel: officials list only
# Defense
rows(DEAD, 855)                                 # Personnel Policy Board: not printed
# Executive Office and independent agencies
rows(DEAD, 746)                                 # Philippine Alien Property Administration: abolished (appendix)
rows(DEAD, 751)                                 # Philippine War Damage Commission: ceased March 31, 1951 (appendix)
rows(DEAD, 648)                                 # War Contracts Price Adjustment Board: abolished March 23, 1951 -> Renegotiation Board
rows(DEAD, 329)                                 # Federal National Mortgage Association: one sentence in the FHA section (p. 413)
rows(DEAD, 875)                                 # National Archives Trust Fund Board: members list and a sentence only (pp. 391, 392)
rows(OTHER, 893)                                # Panama Canal Company: own Independent Agencies heading (p. 441)
rows(GSA, 361)                                  # Federal Records Council: GSA Related Organizations heading (p. 400)

ALSO_1950 = {'Interdepartmental Savings Bond Committee', 'Office of the Deputy Postmaster General',
             'Office of the Executive Assistant to the Postmaster General',
             'Office of Administrative Assistant to the Postmaster General', 'Postal Savings System',
             'Weapons Systems Evaluation Group', 'Scientific Advisory Board (Air Force)', 'Continental Air Command',
             'Strategic Air Command', 'Air Training Command', 'Air Materiel Command', 'Air Proving Ground',
             'Military Air Transport Service', 'Air University', 'Headquarters Command, USAF',
             'Special Weapons Command', 'Air Research and Development Command',
             'United States Air Force Security Service', 'Public Advisory Board (ECA)',
             'Records Management Division (GSA)'}

NEW = [
    ('Office of Defense Mobilization', EOP),                        # p. 66
    # Interior (unit list p. 189; sections pp. 208-211)
    ('Defense Electric Power Administration', INT), ('Defense Fisheries Administration', INT),
    ('Defense Minerals Administration', INT), ('Defense Solid Fuels Administration', INT),
    ('Petroleum Administration for Defense', INT),
    # Commerce
    ('National Production Authority', COM),                         # p. 270; its bureaus and offices not added (depth)
    ('Industry Evaluation Board', COM), ('Business Advisory Council', COM),   # pp. 254-255
    ('Office of Budget and Management (Commerce)', COM), ('Office of Personnel Administration (Commerce)', COM),
    ('Office of Administrative Services (Commerce)', COM), ('Office of Transportation (Commerce)', COM),   # p. 253
    # State
    ('Technical Cooperation Administration', ST),                   # p. 78
    ('Office of International Materials Policy', ST),               # p. 79
    ('Program Planning and Evaluation Staff (State)', ST),          # p. 80
    ('Office of South American Affairs', ARAB),                     # p. 81
    ('Office of Near Eastern Affairs', NEAB), ('Office of African Affairs', NEAB),   # p. 82
    ('International Claims Commission', ST), ('United States National Commission for UNESCO', ST),   # pp. 84-85
    # Treasury
    ('Interdepartmental Savings Bond Committee', TRE),              # Related Organization p. 115 (1950 p. 112)
    # Post Office (pp. 178-179, 184); the same bold headings are in 1950
    ('Office of the Deputy Postmaster General', PO), ('Office of the Executive Assistant to the Postmaster General', PO),
    ('Office of Administrative Assistant to the Postmaster General', PO), ('Postal Savings System', PO),
    # Labor
    ('Defense Manpower Administration', LAB),                       # p. 293
    # Department of Defense (pp. 119-122)
    ('Executive Office of the Secretary (Defense)', NME), ('Office of Management Engineering (Defense)', NME),
    ('Military Traffic Service', NME), ('Weapons Systems Evaluation Group', NME),
    # Air Force: Scientific Advisory Board and the major commands (pp. 157-161; 1950 pp. 152-156)
    ('Scientific Advisory Board (Air Force)', AF), ('Air Defense Command', AF), ('Air Materiel Command', AF),
    ('Air Proving Ground', AF), ('Air Research and Development Command', AF), ('Air Training Command', AF),
    ('Air University', AF), ('Continental Air Command', AF), ('Headquarters Command, USAF', AF),
    ('Military Air Transport Service', AF), ('Special Weapons Command', AF), ('Strategic Air Command', AF),
    ('Tactical Air Command', AF), ('United States Air Force Security Service', AF),
    # Independent agencies
    ('Committee on Purchases of Blind-Made Products', OTHER),       # p. 313
    ('Defense Production Administration', OTHER),                   # p. 317
    ('Defense Transport Administration', OTHER),                    # p. 320
    ('Public Advisory Board (ECA)', ECA),                           # p. 326 (1950 p. 302)
    (ESA, OTHER), ('Salary Stabilization Board', ESA), ('Office of Price Stabilization', ESA),
    ('Wage Stabilization Board', ESA),                              # pp. 326-334
    ('Federal Civil Defense Administration', OTHER),                # p. 335
    ('Office of Defense Coordination (GSA)', GSA), ('Emergency Procurement Service', GSA),   # pp. 393, 396
    ('Records Management Division (GSA)', NARS),                    # run-in p. 399 (1950 p. 357)
    ("President's Advisory Committee on Management", OTHER),        # p. 441
    ('Renegotiation Board', OTHER),                                 # p. 451
    ('Subversive Activities Control Board', OTHER),                 # p. 466
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
    if name in ALSO_1950:
        for i, v in enumerate(vals(c)):
            ws.cell(r, PREV + i).value = v
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
