"""Fill the 1947 columns from the 1947 Government Manual, Second Edition (the latest 1947 edition).

Organization charts (printed pp. 567-603) decide parents first (solid line = authority, dashed = not a
parent), then unit lists, officials pages, the Executive Office (p. 53) and OEM (p. 61) lists, and
Appendix A (the 1946 Reorganization Plans and Executive Order 9809). Every row was run through
skills/gov-manual-agency-entry/scripts/evidence.py. Rows are found by exact name (row numbers in comments).
Page numbers are printed pages (PDF page = printed + 8).
"""
import openpyxl, shutil

SRC = 'data/Spreadsheet 1935-1946.xlsx'
DST = 'data/Spreadsheet 1935-1947.xlsx'
YEAR, PREV_YEAR = 1947, 1946

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, NAVY, TRE, WAR, PO = ('Department of the Interior', 'Department of the Navy', 'Department of the Treasury',
                           'Department of War', 'Post Office Department')
CABINET = {AG, COM, JUS, LAB, ST, INT, NAVY, TRE, WAR, PO}
DEAD, NOTYET, OTHER, CAB = 'DEAD', 'NOTYET', 'OTHER', 'CAB'
EOP, OEM, FSA, SSA = 'Executive Office of the President', 'Office for Emergency Management', 'Federal Security Agency', 'Social Security Administration'
OSS_FSA, SS, BFDC = 'Office of Special Services (FSA)', 'Special Staff', 'Bureau of Foreign and Domestic Commerce'
FL, OIR, OICD = 'Office of Foreign Liquidation', 'Office of Research and Intelligence', 'Office of Intelligence Collection and Dissemination'
EUR, FE, NEA, ARA_ST, PA, FS, CTL = ('Office of European Affairs', 'Office of Far Eastern Affairs', 'Office of Near Eastern and African Affairs',
    'Office of American Republic Affairs', 'Office of Public Affairs', 'Office of the Foreign Service', 'Office of Controls')


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

# Executive Office (p. 53, chart p. 569) and emergency agencies (OEM list p. 61)
rows(EOP, 376)                                  # Office of Government Reports re-established Dec 12, 1946 (FAQ: unhide)
named(DEAD, 'Office of War Mobilization and Reconversion', 'Office of Price Administration', 'Civilian Production Administration',
      'Office of Economic Stabilization', 'Director of Liquidation', 'Committee on Fair Employment Practice',
      'Office of Alien Propety Custodian', 'War Shipping Administration', 'Office of the Coordinator of Inter-American Affairs',
      "President's war Relief Control Board", 'Selective Service System', 'Petroleum Administration for War',
      'Board of War Communications', 'Central Intelligence Group')   # App. A (OTC consolidation Dec 12, 1946 and its termination; other 1946 terminations); CIG only a sentence (p. 78)
named(TRE, 'Office of Contract Settlement')     # transferred to Treasury by E. O. 9809 (p. 151)
# State (chart p. 574; officials pp. 93-100)
named(DEAD, "Secretary's Staff Committee", 'Coordinating Committee (State)', 'Division of Japanese Affairs',
      'Division of Caribbean and Central American Affairs', 'Division of American Republics Analysis and Liaison',
      'Division of Coordination and Review', 'Division of Training Services', 'Division of Research and Publication',
      'Division of British Commonwealth Intelligence', 'Division of USSR Intelligence',
      'Division of Europe, Near East and Africa Intelligence', 'Presentation Division')
named(OICD, 'Division of Map Intelligence and Cartography')      # now "Map Intelligence Division" under OICD
# Treasury (chart p. 575): Procurement Division renamed Bureau of Federal Supply (same row)
# War (chart p. 576): Army Service Forces abolished June 1946; technical and administrative services report to the Chief of Staff
named(DEAD, 'Army Service Forces', 'Strength Accounting and Statistics Office', 'Personnel and Administrative Branch (War)', 'Secretariat (War)')
named(WAR, 'Office of The Adjutant General', 'Office of the Chief of Chaplains',
      'Office of the Chief of Finance', 'Office of the Chief of Engineers', 'Office of the Chief Signal Officer',
      'Office of the Quartermaster General', 'Office of the Surgeon General', 'Office of the Chief of Ordnance',
      'Office of the Chief of the Chemical Warfare Service', 'Office of the Provost Marshal General', 'Office of the Chief of Transportation')
rows(WAR, 276)                                  # (War JAG row; the Navy JAG row 231 is separate)
# Post Office (list p. 189): Second Assistant's bureau reorganized into postal transport divisions
named(DEAD, 'Division of Railway Mail Service', 'Division of Railway Adjustments', 'Division of Protocol')  # replaced by Surface Postal Transport; State protocol now a staff of the Executive Secretariat
# Justice (chart p. 577): U. S. Attorneys and Marshals now on the main line
named(JUS, 'United States Attorneys', 'United States Marshals')
# Navy (chart p. 579): Office of Judge Advocate General is a box (FAQ: unhide row 231)
rows(NAVY, 231)
# Interior (chart p. 580; list p. 240)
named(DEAD, 'General Land Office', 'Grazing Service', 'Petroleum Conservation Division', 'War Relocation Authority')  # BLM consolidation July 16, 1946; Oil and Gas Division; WRA terminated June 30, 1946
rows(INT, 179)                                  # Coal Mines Administration reappears in list and chart (FAQ: unhide)
# Agriculture (chart p. 581)
named(DEAD, 'Farm Security Administration', 'Emergency Crop and Feed Loans', 'Regional Agricultural Credit Corporations')  # FSA abolished Aug 14, 1946 -> Farmers Home Administration; FCA programs only run-in/sentences
# Commerce (chart p. 582): Civil Aeronautics Board joined by a dashed line only
named(OTHER, 'Civil Aeronautics Board')
# Labor (chart p. 583)
named(DEAD, 'Retraining and Reemployment Administration', 'National Wage Stabilization Board', 'Wage Adjustment Board',
      'Shipbuilding Stabilization Committee')
named(SSA, "Children's Bureau")                 # transferred to FSA July 16, 1946; inside the Social Security Administration box (p. 585)
# Federal Security Agency (chart p. 585)
named(DEAD, 'Social Security Board', 'Bureau of Research and Statistics', 'Bureau of Accounts and Audtis',
      'Office of Community War Services', 'Office of War Property Distribution', 'United States Employees Compensation Commission',
      'Office of Economic Research (FWA)')
named(SSA, 'Bureau of Old-Age and Survivors Insurance', 'Bureau of Employment Security', 'Bureau of Public Assistance')
named(OSS_FSA, 'Food and Drug Administration', 'Office of Vocational Rehabilitation')

NEW = [
    ('Council of Economic Advisers', EOP),                          # p. 58, chart p. 569
    ('Philippine Alien Property Administration', OEM),              # p. 63
    ('Advisory Committee on Voluntary Foreign Aid', OTHER),         # p. 66
    ('Office of the Housing Expediter', OTHER),                     # p. 79
    ('Office of Selective Service Records', OTHER),                 # p. 80
    ('Price Decontrol Board', OTHER),                               # p. 81
    ('Philippine War Damage Commission', OTHER),                    # p. 87
    ('Atomic Energy Commission', OTHER),                            # p. 396
    ('Indian Claims Commission', OTHER),                            # p. 438
    # State
    ('Executive Secretariat (State)', ST),                          # chart p. 574
    ('Division of Northeast Asian Affairs', FE),
    ('Division of Caribbean Affairs', ARA_ST),
    ('Division of Central America and Panama Affairs', ARA_ST),
    ('Division of Special Inter-American Affairs', ARA_ST),
    ('Division of Public Studies', PA), ('Division of Historical Policy Research', PA), ('Division of Publications', PA),
    ('Foreign Service Institute', FS),
    ('Munitions Division', CTL),
    ('Division of Research for Europe', OIR), ('Division of Research for Near East and Africa', OIR),
    ('Administrative Division (Foreign Liquidation)', FL), ('Fiscal and Accounting Division (Foreign Liquidation)', FL),
    ('General Disposals Division (Foreign Liquidation)', FL), ('Legal Division (Foreign Liquidation)', FL),
    ('Maritime Division (Foreign Liquidation)', FL), ('Research and Statistics Division (Foreign Liquidation)', FL),
    # Post Office (list p. 189)
    ('Division of Surface Postal Transport', 'Bureau of the Second Assistant Postmaster General'),
    ('Division of Administrative Services (Post Office)', 'Bureau of the Second Assistant Postmaster General'),
    # Justice
    ('Office of Alien Property', JUS),                              # chart p. 577
    # War Special Staff boxes (chart p. 576)
    ('Legislative and Liaison Division', SS), ('Troop Information and Education Division', SS),
    ('Historical Division (War)', SS), ('War Department Manpower Board', SS), ('Budget Division (War)', SS),
    ('Civil Affairs Division', SS),
    ('Special Services (War)', WAR),                                # administrative staffs and services
    # Interior (chart p. 580)
    ('Bureau of Land Management', INT), ('Oil and Gas Division', INT), ('Division of Geography (Interior)', INT),
    ('Office of the Safety Engineer (Interior)', INT),
    # Agriculture (chart p. 581)
    ('Farmers Home Administration', AG), ('Sugar Rationing Administration', AG),
    ('Commodity Exchange Authority', AG), ('Office of Hearing Examiners (Agriculture)', AG),
    # Commerce (chart p. 582)
    ('Office of Technical Services', COM),
    # Federal Security Agency (chart p. 585)
    (SSA, FSA), ('Office of Appeals Council', SSA), ("Freedmen's Hospital", FSA),
    (OSS_FSA, FSA), ("Bureau of Employees' Compensation", OSS_FSA), ("Employees' Compensation Appeals Board", OSS_FSA),
    ('Office of Administration (FSA)', FSA), ('Office of Research (FSA)', FSA), ('Office of General Counsel (FSA)', FSA),
    ('Office of Information (FSA)', FSA), ('Office of Federal-State Relations', FSA),
    ('Office of Inter-Agency and International Relations', FSA),
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
