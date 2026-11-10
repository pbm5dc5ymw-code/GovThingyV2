"""Fill the 1955 columns from the United States Government Organization Manual 1955-56 (revised as of June 1, 1955;
the only edition). Run after fill1954.py. Page numbers are printed pages: PDF = printed + 8 throughout.

Same three requirements as 1944-54: an agency name (not a position), set apart by its formatting, followed by a
description. Every row live in 1954 was matched to its 1955 heading; changed units were read on the page images.

1955 changes that matter:
  - Treasury: a bold Office of Production and Defense Lending (December 7, 1954) takes over the liquidation of the
    Reconstruction Finance Corporation; the RFC (liquidating), the Federal Facilities Corporation, and the Defense
    Lending Division are run-ins under it (pp. 98-99).
  - Commerce: the staff offices are small-caps run-ins with descriptions again (pp. 258-260), after plain lists in
    1952-54, so Budget and Management, Publications, Personnel Management, and Administrative Operations come back
    on their old rows, and Public Information, Strategic Information, International Trade Fairs, Security Control,
    the Agency Inspection Staff, the Advisory Committee on Export Policy, and the Defense Air Transportation
    Administration are new.
  - Interior: "Other Departmental Units" (pp. 206-209) has bold Offices of the Solicitor, Geography, Minerals
    Mobilization, and Oil and Gas, a Saline Water Conversion Program Staff, and a Division of Information. The
    Defense Solid Fuels Administration and the Puerto Rico Reconstruction Administration are not printed.
  - State: the Bureau of United Nations Affairs is now the Bureau of International Organization Affairs (p. 83);
    under the Controller the Office of Communications and Records becomes part of a new Office of Special Services,
    and the Foreign Service Institute and Inspection Corps are run-ins again (pp. 81, 86-87).
  - Post Office: the Chief Post Office Inspector is now only a title heading and the Bureau of the Controller is
    gone; the Bureau of Finance covers its work (pp. 198-200).
  - Justice: new Internal Security Division (July 1954, p. 190). HEW: new run-in Offices of Inspection and of the
    Comptroller (p. 310).
  - Independent agencies: the Housing and Home Finance Agency has new bold Community Facilities and Urban Renewal
    Administrations (December 1954, p. 421); the National Housing Council is only named in a sentence (p. 420). The
    Foreign Operations Administration has a new Foreign Operations Council and Education Committee; its Public
    Advisory Board is gone (pp. 396-400). New Civil Defense Coordinating Board (May 1955, p. 343) and President's
    Committee on Government Employment Policy (January 1955, p. 457), which replaces the Fair Employment Board. The
    War Claims Commission is now part of the Foreign Claims Settlement Commission (p. 395); the Defense Transport
    Administration is not printed. The President's Advisory Committee on Government Organization is printed in the
    Executive Office (p. 74).
Renames keep the same row: Army Ground Forces -> Continental Army Command (286); Bureau of United Nations Affairs ->
Bureau of International Organization Affairs (903); Office of United Nations Economic and Social Affairs -> Office
of International Economic and Social Affairs (919); Office of Financial and Development Policy -> Office of
International Financial and Development Affairs (654); Office of International Trade Policy -> Office of
International Trade and Resources (653, which absorbs International Materials Policy); Office of Transportation and
Communications -> Office of Transport and Communications (577); Office of Budget and Reports -> Office of Budget
(1047); Division of Cryptography -> Cryptography Staff (1060); Division of Foreign Reporting Services -> Foreign
Reporting Staff (675); Division of Foreign Buildings Operations -> Office of Foreign Buildings (625); Office of the
Solicitor (Agriculture) -> Office of the General Counsel (637); Office of Personnel Administration (Commerce) ->
Office of Personnel Management (966); Office of Administrative Services (Commerce) -> Office of Administrative
Operations (967); Oil and Gas Division -> Office of Oil and Gas (783); Labor and Health Advisory Committees (FOA) ->
Labor and Health Committees (1081, 1082).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1954.xlsx'
DST = 'data/Spreadsheet 1935-1955.xlsx'
YEAR, PREV_YEAR = 1955, 1954

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
HEW = 'Department of Health, Education, and Welfare'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME, HEW}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
EOP, FOA, HHFA = 'Executive Office of the President', 'Foreign Operations Administration', 'Housing and Home Finance Agency'
OPDL = 'Office of Production and Defense Lending'


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
rows(EOP, 1058)                                 # President's Advisory Committee on Government Organization (p. 74)
# State (pp. 81-89)
rows(ST, 762, 1021)                             # Foreign Service Institute, Inspection Corps: run-ins again (p. 81)
rows(DEAD, 970)                                 # Office of International Materials Policy: in Trade and Resources
rows(DEAD, 1061)                                # Office of Communications and Records -> Office of Special Services
rows(DEAD, 975)                                 # International Claims Commission: now in the Foreign Claims
                                                # Settlement Commission (p. 395)
# Treasury
rows(OPDL, 432)                                 # Reconstruction Finance Corporation (liquidating) run-in (p. 99)
# Post Office (pp. 198-200)
rows(DEAD, 422)                                 # Chief Post Office Inspector: title heading only
rows(DEAD, 424)                                 # Bureau of the Controller: not printed
# Interior (pp. 206-209)
rows(INT, 186, 199, 783)                        # Division of Information, Office of the Solicitor, Office of Oil and Gas
rows(DEAD, 960, 207)                            # Defense Solid Fuels Administration, Puerto Rico Reconstruction
                                                # Administration: not printed
# Commerce (p. 260)
rows(COM, 965, 932, 966, 967)                   # Budget and Management, Publications, Personnel Management,
                                                # Administrative Operations: run-ins again
# Independent agencies
rows(DEAD, 880)                                 # National Housing Council: named in a sentence only (p. 420)
rows(DEAD, 1004)                                # Public Advisory Board (FOA): not printed
rows(DEAD, 853)                                 # War Claims Commission: in the Foreign Claims Settlement Commission
rows(DEAD, 1003)                                # Defense Transport Administration: not printed
rows(DEAD, 894)                                 # Fair Employment Board -> President's Committee on Government
                                                # Employment Policy

NEW = [
    # State (pp. 87)
    ('Office of Special Services (State)', ST), ('Regulations and Procedures Staff (State)', ST),
    # Treasury (pp. 98-99)
    (OPDL, TRE), ('Federal Facilities Corporation', OPDL), ('Defense Lending Division', OPDL),
    # Justice (p. 190)
    ('Internal Security Division', JUS),
    # Interior (pp. 207-209)
    ('Office of Geography', INT), ('Office of Minerals Mobilization', INT),
    ('Saline Water Conversion Program Staff', INT),
    # Commerce (pp. 258-260)
    ('Office of Public Information (Commerce)', COM), ('Office of Strategic Information', COM),
    ('Defense Air Transportation Administration', COM), ('Office of International Trade Fairs', COM),
    ('Advisory Committee on Export Policy', COM), ('Office of Security Control (Commerce)', COM),
    ('Agency Inspection Staff (Commerce)', COM),
    # HEW (p. 310)
    ('Office of Inspection (HEW)', HEW), ('Office of the Comptroller (HEW)', HEW),
    # Independent agencies
    ('Civil Defense Coordinating Board', OTHER),                                     # p. 343
    ('Foreign Operations Council', FOA), ('Education Committee (FOA)', FOA),          # pp. 399-400
    ('Community Facilities Administration', HHFA), ('Urban Renewal Administration', HHFA),   # p. 421
    ("President's Committee on Government Employment Policy", OTHER),               # p. 457
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
