"""Fill the 1945 columns from the 1945 Government Manual, Second Edition (personnel as of September 20, 1945).

The second edition is the most recent 1945 edition (FAQ rule). It "pictures the activities of [the war]
agencies as they functioned just prior to V-J day", so footnoted September 1945 terminations are flagged
but the printed structure is coded. No organization charts: structure comes from page images of each
department's officials page and unit list, the Executive Office (p. 53) and OEM (p. 60) membership lists,
and Appendix A. Every row was also run through skills/gov-manual-agency-entry/scripts/evidence.py; a row
stays live only with printed evidence (heading, subheading, list entry or officials line).
Page numbers are printed pages (PDF page = printed + 8).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1944.xlsx'
DST = 'data/Spreadsheet 1935-1945.xlsx'
YEAR, PREV_YEAR = 1945, 1944

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, NAVY, TRE, WAR, PO = ('Department of the Interior', 'Department of the Navy', 'Department of the Treasury',
                           'Department of War', 'Post Office Department')
CABINET = {AG, COM, JUS, LAB, ST, INT, NAVY, TRE, WAR, PO}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
OWMR, FLA, FWA = 'Office of War Mobilization and Reconversion', 'Federal Loan Agency', 'Federal Works Agency'
# State offices after the December 20, 1944 reorganization (pp. 181-184)
ITP, FDP, TCP = 'Office of International Trade Policy', 'Office of Financial and Development Policy', 'Office of Transportation and Communications'
PA, IICA, IIIS = 'Office of Public Affairs', 'Office of International Information and Cultural Affairs', 'Interim International Information Service'
SPA, ARA_ST, EUR, NEA, FE = ('Office of Special Political Affairs', 'Office of American Republic Affairs', 'Office of European Affairs',
                             'Office of Near Eastern and African Affairs', 'Office of Far Eastern Affairs')
CTL, DA, FS = 'Office of Controls', 'Office of Departmental Administration', 'Office of the Foreign Service'


def vals(c):
    """Five-column coding: cabinet, child, cabinet parent, other parent, other agency."""
    if c == DEAD: return ['Dead'] * 5
    if c == NOTYET: return ['No', 'No', 'None', 'None', 'No']
    if c == CAB: return ['Yes', 'No', 'None', 'None', 'No']
    if c == OTHER: return ['No', 'No', 'None', 'None', 'Yes']
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


O = {}
def rows(code, *row_numbers):
    O.update({r: code for r in row_numbers})

# Executive Office (p. 53) and emergency war agencies
rows(DEAD, 378)                                 # Committee for Congested Production Areas terminated Dec 31, 1944 (App. A); not in EOP list
rows(DEAD, 379)                                 # Division of Central Administrative Services terminated Nov 30, 1944 (App. A)
rows(DEAD, 488, 509)                            # Office of Civilian Defense abolished June 30, 1945 (App. A); not in OEM list (p. 60)
rows(DEAD, 495)                                 # Office of War Mobilization: functions transferred to OWMR Oct 3, 1944 (p. 141)
rows(DEAD, 563)                                 # Surplus War Property Administration succeeded by Surplus Property Board (p. 144)
rows(OWMR, 564, 565)                            # Retraining and Reemployment Administration, Office of Contract Settlement placed in OWMR (p. 141, 143)
rows(DEAD, 16)                                  # Board of Investigation and Research ceased Sept 18, 1944 (App. A)
rows(OTHER, 304)                                # Export-Import Bank: permanent independent agency July 31, 1945; own section (p. 451)
# Federal Loan Agency re-created Feb 24, 1945 (p. 403): FAQ - unhide and relabel the old row
rows(OTHER, 328)
rows(FLA, 432)                                  # Reconstruction Finance Corporation (p. 403)
rows(DEAD, 433, 434, 435, 436)                  # Defense Plant, Rubber Reserve, Metals Reserve, Defense Supplies dissolved June 30, 1945 into RFC (App. A; pp. 407-408 past tense)
rows(DEAD, 297)                                 # Disaster Loan Corporation dissolved June 30, 1945 (App. A)
# State: offices and divisions not in the 1945 officials list (pp. 181-183)
rows(DEAD, 578, 579, 580, 583, 584, 585)        # Wartime Economic Affairs, Economic Affairs, Public Information; Policy and Post-War Programs committees (replaced by Secretary's Staff and Coordinating Committees); ECEFP only in a sentence (p. 185)
rows(DEAD, 586, 587, 590, 592, 602, 608, 609, 610, 611, 613, 616, 618, 619, 623, 123, 150)
rows(DEAD, 22)                                  # Committee for Reciprocity Information: only sentences (pp. 560, 597); no section
rows(ITP, 612, 614, 615)                        # Commercial Policy, Commodities, Petroleum (p. 182)
rows(PA, 617, 147, 624)                         # Public Liaison, Research and Publication, Geography and Cartography (p. 182)
rows(IICA, 620)                                 # Central Translating Division (p. 182)
# Agriculture (p. 333): War Food Administration terminated June 30, 1945 (App. A)
rows(DEAD, 550, 551, 552, 627, 628, 629, 630)
rows(DEAD, 4)                                   # Agricultural Adjustment Agency consolidated into Production and Marketing Administration Aug 18, 1945 (App. A)
rows(AG, 24, 45, 626, 306, 54)                  # CCC, FCIC, Extension Service, Farm Security Administration, Soil Conservation Service report to the Secretary (p. 333)
rows(DEAD, 319)                                 # Federal Farm Mortgage Corporation: only sentences in FCA text (pp. 345, 347); no subheading
# Interior (p. 309)
rows(DEAD, 179)                                 # Coal Mines Administration not in the unit list
# War (pp. 233-235)
rows(WAR, 294, 539)                             # National Guard Bureau, Executive for Reserve and ROTC Affairs now listed above Army Service Forces (p. 234)
rows(DEAD, 537)                                 # Armored Center not in Army Ground Forces list (p. 235)
# Federal Works Agency (p. 444)
rows(DEAD, 546)                                 # War Public Works program now administered by the Bureau of Community Facilities

NEW = [
    ('Office of War Mobilization and Reconversion', OTHER),        # p. 140-143; statute of Oct 3, 1944; not in OEM list
    ('Surplus Property Board', OWMR),                               # p. 144
    ('War Contracts Price Adjustment Board', OTHER),                # p. 156-157
    ('Office of Civilian Defense Property', COM),                   # Commerce list p. 368-369
    ('Office of Surplus Property', COM),                            # Commerce list p. 369
    ('Production and Marketing Administration', AG),                # p. 333
    ('Bureau of Community Facilities', FWA),                        # p. 444
    # State offices (p. 184) and their divisions (pp. 181-183)
    (ITP, ST), (FDP, ST), (PA, ST), (IICA, ST), (IIIS, ST),
    ("Secretary's Staff Committee", ST),                            # p. 181
    ('Coordinating Committee (State)', ST),                         # p. 181
    ('Division of Southeast Asian Affairs', FE),
    ('Division of North and West Coast Affairs', ARA_ST),
    ('Division of International Organization Affairs', SPA),
    ('Division of International Security Affairs', SPA),
    ('Division of Dependent Area Affairs', SPA),
    ('War Areas Economic Division', ITP),
    ('Division of International Labor, Social, and Health Affairs', ITP),
    ('Division of Financial Affairs', FDP),
    ('Division of Foreign Economic Development', FDP),
    ('Division of Economic Security Controls', FDP),
    ('Division of Lend-Lease and Surplus War Property Affairs', FDP),
    ('Division of Cultural Cooperation', IICA),
    ('International Information Division', IICA),
    ('Division of Foreign Service Planning', FS),
    ('Division of Training Services', FS),
    ('Division of Foreign Reporting Services', FS),
    ('Division of Management Planning', DA),
    ('Division of Central Services', DA),
    ('Division of Cryptography', DA),
]

shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
COL, PREV = hdr[f'{YEAR}_cabinet_agency'], hdr[f'{PREV_YEAR}_cabinet_agency']

last = max(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value)
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
