"""Fill the 1949 columns from the 1949 United States Government Organization Manual (only 1949 edition).

Organization charts (printed pp. 571-612) decide parents first, then unit lists, officials pages, the
Executive Office list (p. 53), the OEM paragraph (p. 64), and Appendix A. Every row was run through
evidence.py. Page numbers are printed pages (PDF page = printed + 10).

The text reflects the General Services Administration, established July 1, 1949 (section p. 417): the
Federal Works Agency and its units, the Bureau of Federal Supply, Contract Settlement, the National
Archives Establishment and its related bodies, and War Assets functions were transferred to GSA. Older
sections are still printed with footnotes "Transferred on July 1, 1949, to General Services
Administration"; the GSA section is coded as the current structure (flagged for the user).
The National Military Establishment convention from 1948 is kept (NME = cabinet department).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1948.xlsx'
DST = 'data/Spreadsheet 1935-1949.xlsx'
YEAR, PREV_YEAR = 1949, 1948

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, TRE, PO = 'Department of the Interior', 'Department of the Treasury', 'Post Office Department'
NME = 'National Military Establishment'
CABINET = {AG, COM, JUS, LAB, ST, INT, TRE, PO, NME}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
GSA, SSA, BFDC, BES = 'General Services Administration', 'Social Security Administration', 'Bureau of Foreign and Domestic Commerce', 'Bureau of Employment Security'
B1 = 'Bureau of the First Assistant Postmaster General'
EUR, UNA = 'Office of European Affairs', 'Office of United Nations Affairs'
GAA, CA, OF, PER = 'Office of German and Austrian Affairs', 'Office of Consular Affairs', 'Office of Operating Facilities', 'Office of Personnel (State)'
BF = 'Office of Budget and Finance (State)'


def vals(c):
    """Five-column coding: cabinet, child, cabinet parent, other parent, other agency."""
    if c == DEAD: return ['Dead'] * 5
    if c == NOTYET: return ['No', 'No', 'None', 'None', 'No']
    if c == CAB: return ['Yes', 'No', 'None', 'None', 'No']
    if c == OTHER: return ['No', 'No', 'None', 'None', 'Yes']
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


O = {}
NAMED = []
def rows(code, *row_numbers):
    O.update({r: code for r in row_numbers})
def named(code, *names):
    NAMED.extend((n, code) for n in names)

# Executive Office (p. 53, chart p. 573); OEM now consists of the Philippine Alien Property Administration (p. 64)
rows(DEAD, 492)                                  # Office of Defense Transportation: historical mention only
rows(OTHER, 22)                                  # Committee for Reciprocity Information re-created Oct 5, 1948; own section p. 338 (FAQ: unhide)
# General Services Administration (p. 417)
rows(DEAD, 343, 709, 710, 712, 681)              # Federal Works Agency and its staff offices; War Assets Administration abolished June 30, 1949 (App. A)
rows(GSA, 429, 430, 652, 320, 332)               # Public Buildings, Public Roads, Community Facilities, Federal Fire Council, Federal Real Estate Board
rows(GSA, 264, 565)                              # Bureau of Federal Supply (Procurement Division row), Office of Contract Settlement
rows(GSA, 360, 361, 366)                         # National Archives Establishment and related bodies
# National Military Establishment (chart p. 580)
rows(DEAD, 812, 813)                             # Armed Services Petroleum Board, Army-Navy Explosives Safety Board: not printed
# State (chart p. 578; officials pp. 83-94)
rows(DEAD, 576, 581, 582, 135, 137, 673, 835, 606, 596, 597, 714, 766, 767, 768, 769, 770, 771, 824, 833)
rows(CA, 166, 170, 836, 763, 723)                # Passport, Visa, Protective Services, Munitions, Security (Investigations) divisions
rows(OF, 123, 677, 678, 620)                     # Communications and Records, Central Services, Cryptography, Language Services
rows(PER, 129, 138)                              # Departmental Personnel, Foreign Service Personnel
rows(UNA, 140)                                   # Division of International Conferences
rows(ST, 625, 675, 762)                          # Foreign Buildings Operations, Foreign Reporting Services, Foreign Service Institute (Director General is an official)
rows(BF, 720)                                    # Division of Budget reappears under Office of Management and Budget (FAQ: unhide)
# Interior (chart p. 586)
rows(INT, 211, 173)                              # Alaska Railroad, Alaska Road Commission on chart (FAQ: unhide)
# Commerce (chart p. 588)
rows(BFDC, 702)                                  # Office of Small Business back on chart (FAQ: unhide)
# Labor / Federal Security (charts pp. 589, 596): U. S. Employment Service transferred to FSA July 1, 1948
rows(BES, 105)
rows(OTHER, 848, 849)                            # Howard University, Columbia Institution: "Federally Aided Corporations" joined by a DASHED line
rows('Bureau of Medical Services', 793)          # Freedmen's Hospital indented under Bureau of Medical Services in the PHS box
# Post Office (list pp. 208-209)
rows(DEAD, 390)                                  # Division of Post Office Service replaced by new First Assistant divisions

NEW = [
    (GSA, OTHER),                                                   # p. 417
    ('Displaced Persons Commission', OTHER),                        # p. 65
    ('Motor Carrier Claims Commission', OTHER),                     # p. 71
    ('War Claims Commission', OTHER),                               # p. 76
    ('Military Liaison Committee to the Atomic Energy Commission', NME),   # chart p. 580, officials p. 141
    ('Personnel Policy Board', NME), ('Civilian Components Policy Board', NME),
    ('Tax Advisory Staff', TRE),                                    # Treasury list p. 107
    ('Virgin Islands Company', INT),                                # chart p. 586
    ('Bureau of Federal Credit Unions', SSA),                       # chart p. 596
    (GAA, ST), (CA, ST), (OF, ST), (PER, ST),
    ('Division of German Political Affairs', GAA), ('Division of German Economic Affairs', GAA),
    ('Division of Austrian Affairs', GAA), ('Division of German and Austrian Information and Reorientation Affairs', GAA),
    ('Division of Southwest European Affairs', EUR), ('Division of Southeast European Affairs', EUR),
    ('Division of Budget and Administrative Services (Post Office)', B1), ('Division of Post Office Clerical Service', B1),
    ('Division of City Delivery Service', B1), ('Division of Post Office Personnel', B1), ('Division of Rural Delivery Service', B1),
]


shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
COL, PREV = hdr[f'{YEAR}_cabinet_agency'], hdr[f'{PREV_YEAR}_cabinet_agency']

last = max(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value)
byname = {}
for r in range(3, last + 1):
    if ws.cell(r, 1).value:
        byname.setdefault(str(ws.cell(r, 1).value).strip(), []).append(r)
for n, c in NAMED:
    assert n in byname, n
    for r in byname[n]:
        O[r] = c
for r in range(3, last + 1):
    if not ws.cell(r, 1).value:
        continue
    prev = [ws.cell(r, PREV + i).value for i in range(5)]
    if r in O:
        new = vals(O[r])
        if O[r] not in (DEAD, NOTYET) and not (ws.cell(r, 2).value or '').strip():
            ws.cell(r, 2).value = 'Executive'
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
        ws.row_dimensions[x].hidden = kind in ('Legislative', 'Judicial') or ws.cell(x, COL).value == 'Dead'

wb.save(DST)
print('saved', DST, 'new rows', last + 1, '-', r - 1)
