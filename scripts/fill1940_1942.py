"""Fill the 1940, 1941 and 1942 columns of the agency spreadsheet.

Starts from the user's corrected 1939 sheet and applies changes found in the
U.S. Government Manuals of September 1940, September 1941 and March 1942.
Every row carries its previous year's coding forward unless an override below
says otherwise; Dead rows stay Dead.
"""
import openpyxl, shutil

SRC = 'data/1939 Agencies!.xlsx'
DST = 'data/Spreadsheet 1935-1942.xlsx'
YEARS = {1940: 23, 1941: 28, 1942: 33}  # first column of each year's block

AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, NAVY, TRE, WAR, PO = ('Department of the Interior', 'Department of the Navy', 'Department of the Treasury',
                           'Department of War', 'Post Office Department')
CABINET = {AG, COM, JUS, LAB, ST, INT, NAVY, TRE, WAR, PO}
FSA, FWA, OEM, SSB = 'Federal Security Agency', 'Federal Works Agency', 'Office for Emergency Management', 'Social Security Board'
CND, CSC, OSRD = 'United States Council of National Defense', 'United States Civil Service Commission', 'Office of Scientific Research and Development'
RFC, FISCAL = 'Reconstruction Finance Corporation', 'Fiscal Service'

DEAD, NOTYET, OTHER, CAB, SKIP = 'DEAD', 'NOTYET', 'OTHER', 'CAB', 'SKIP'


def vals(c):
    if c == DEAD: return ['Dead'] * 5
    if c == NOTYET: return ['No', 'No', 'None', 'None', 'No']
    if c == CAB: return ['Yes', 'No', 'None', 'None', 'No']
    if c == OTHER: return ['No', 'No', 'None', 'None', 'Yes']
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


# row -> {first year the change applies: new coding}; carries forward to later years
O = {
    # Executive Office / emergency agencies
    381: {1940: OTHER},                                     # Office for Emergency Management (May 1940)
    379: {1940: NOTYET, 1941: OEM},                         # Division of Central Administrative Services
    26: {1940: CSC},                                        # Council of Personnel Administration: CSC division list
    462: {1940: OTHER},                                     # Council of National Defense, reestablished May 1940
    516: {1940: WAR, 1941: OTHER},                          # Selective Service: under War in 1940 defense section
    490: {1940: OTHER, 1941: OEM},                          # Defense Communications Board
    489: {1941: OEM, 1942: DEAD},                           # Division of Defense Aid Reports -> Lend-Lease Admin
    509: {1941: OEM},                                       # Office of Civilian Defense
    491: {1941: OEM},                                       # Office of Defense Health and Welfare Services
    500: {1941: OEM},                                       # Division of Defense Housing Coordination
    501: {1941: OEM, 1942: DEAD},                           # National Defense Mediation Board -> NWLB
    502: {1941: OEM, 1942: DEAD},                           # Office of Production Management -> WPB
    503: {1941: OEM},                                       # Office of Scientific Research and Development
    504: {1941: AG},                                        # Office of Agricultural Defense Relations
    505: {1941: OTHER, 1942: DEAD},                         # Economic Defense Board -> Board of Economic Warfare
    511: {1941: OEM},                                       # Office of Price Administration
    493: {1942: OEM},                                       # Office of Lend-Lease Administration
    507: {1942: OEM},                                       # National War Labor Board
    497: {1942: OEM},                                       # War Production Board
    492: {1942: OEM},                                       # Office of Defense Transportation
    514: {1942: OTHER},                                     # Office of Censorship
    512: {1942: OTHER},                                     # Board of Economic Warfare
    433: {1940: RFC}, 434: {1940: RFC}, 435: {1940: RFC}, 436: {1940: RFC}, 437: {1940: RFC},
    71: {1940: COM},                                        # National Inventors Council
    16: {1940: NOTYET, 1942: OTHER},                        # Board of Investigation and Research-Transportation
    # Agriculture
    29: {1940: AG}, 56: {1940: AG}, 57: {1940: AG},         # AMS, Foreign Agricultural Relations, Surplus Marketing Admin
    46: {1940: DEAD},                                       # FDA moved to FSA (row 478 carries it)
    48: {1940: DEAD}, 338: {1940: DEAD},                    # merged into Surplus Marketing Administration
    52: {1940: COM},                                        # Weather Bureau -> Commerce
    # Commerce
    18: {1940: COM}, 19: {1940: COM},                       # Civil Aeronautics Administration / Board
    17: {1940: NOTYET},                                     # 1939 values were shifted in the key
    # Justice and Labor
    85: {1940: JUS}, 102: {1940: DEAD},                     # Immigration and Naturalization -> Justice
    90: {1940: JUS},                                        # Board of Parole
    # State
    117: {1940: DEAD}, 148: {1940: DEAD},
    116: {1940: ST}, 119: {1940: ST}, 121: {1940: ST}, 156: {1940: ST},
    122: {1940: ST, 1942: DEAD}, 124: {1942: DEAD},
    135: {1941: ST}, 150: {1941: ST}, 151: {1941: ST},
    120: {1940: NOTYET, 1942: ST},
    128: {1942: ST}, 130: {1942: ST}, 131: {1942: ST}, 134: {1942: ST}, 136: {1942: ST}, 155: {1942: ST}, 165: {1942: ST},
    # Interior
    62: {1940: DEAD}, 176: {1940: DEAD}, 190: {1940: INT},  # Fish and Wildlife Service
    174: {1940: DEAD}, 175: {1940: INT},                    # Bonneville Power Administration
    180: {1941: DEAD}, 181: {1941: DEAD}, 210: {1941: DEAD},
    183: {1941: INT}, 201: {1941: INT}, 198: {1941: INT},
    193: {1940: FSA},                                       # Office of Education
    # Navy
    214: {1940: NAVY}, 226: {1940: NAVY},
    219: {1940: DEAD}, 220: {1940: DEAD}, 229: {1940: DEAD},
    217: {1942: NAVY},
    # Treasury
    245: {1940: FSA},                                       # Public Health Service
    253: {1940: TRE}, 235: {1940: FISCAL}, 243: {1940: FISCAL}, 261: {1940: FISCAL},
    255: {1940: DEAD}, 265: {1940: DEAD}, 254: {1940: DEAD}, 247: {1940: DEAD},
    251: {1940: TRE},
    267: {1942: NAVY},                                      # Coast Guard -> Navy, Nov 1941
    246: {1942: DEAD}, 242: {1940: NOTYET, 1942: TRE},     # Enrollment and Disbarment -> Committee on Practice
    # War
    273: {1941: WAR}, 283: {1940: NOTYET, 1941: WAR}, 281: {1940: NOTYET, 1941: WAR},
    # Misc
    475: {1941: DEAD},                                      # Washington National Monument Society
    332: {1940: FWA},                                       # Federal Real Estate Board (fixes "Federal Works System")
    484: {1940: SSB}, 485: {1940: SSB}, 486: {1940: SSB}, 487: {1940: SSB},
    # Quasi-official agencies in 1941-42 manuals: not considered
    6: {1941: SKIP}, 358: {1941: SKIP}, 383: {1941: SKIP},
}

NEW = [
    ('Advisory Commission to the Council of National Defense', {1940: CND, 1941: DEAD}),
    ('National Defense Research Committee', {1940: CND, 1941: OSRD}),
    ('Office for Coordination of Commercial and Cultural Relations Between the American Republics', {1940: CND, 1941: DEAD}),
    ('Office of the Coordinator of Inter-American Affairs', {1941: OEM}),
    ('Priorities Board', {1940: OTHER, 1941: DEAD}),
    ('Office of Export Control', {1940: OTHER, 1941: 'Economic Defense Board', 1942: 'Board of Economic Warfare'}),
    ('Interdepartmental Committee for Coordination of Foreign and Domestic Military Purchases', {1940: OTHER, 1941: DEAD}),
    ('Supply Priorities and Allocations Board', {1941: OEM, 1942: DEAD}),
    ('Coordinator of Information', {1941: OTHER}),
    ('Office of Facts and Figures', {1942: OEM}),
    ('Defense Savings Staff', {1942: TRE}),
    ('Foreign Funds Control', {1942: TRE}),
    ('Office of Solid Fuels Coordinator for National Defense', {1942: INT}),
    ('Division of Commercial Policy and Agreements', {1942: ST}),
    ('Division of Exports and Defense Aid', {1942: ST}),
    ('Division of Studies and Statistics', {1942: ST}),
    ('Office of the Provost Marshal General', {1942: WAR}),
    ('Office of the Chief of the Armored Force', {1942: WAR}),
    ('General Headquarters', {1942: WAR}),
    ('Office of the Executive for Reserve and R.O.T.C. Affairs', {1942: WAR}),
]

shutil.copy(SRC, DST)
wb = openpyxl.load_workbook(DST)
ws = wb.worksheets[0]
for y, c in YEARS.items():
    assert ws.cell(2, c).value == f'{y}_cabinet_agency'


def write(r, col, v):
    for i, x in enumerate(v):
        ws.cell(r, col + i).value = x


last = max(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value)
for r in range(3, last + 1):
    if not ws.cell(r, 1).value:
        continue
    prev = [ws.cell(r, c).value for c in range(18, 23)]  # 1939 values
    state = None if all(v in (None, '') for v in prev) else prev
    for y, col in YEARS.items():
        ov = O.get(r, {}).get(y)
        if ov is not None:
            state = None if ov == SKIP else vals(ov)
            if ov not in (DEAD, NOTYET, SKIP) and not (ws.cell(r, 2).value or '').strip():
                ws.cell(r, 2).value = 'Executive'
        elif state and str(state[0]).strip() == 'Dead':
            state = vals(DEAD)
        if state:
            write(r, col, state)

r = last + 1
for name, changes in NEW:
    ws.cell(r, 1).value = name
    ws.cell(r, 2).value = 'Executive'
    for col in range(3, 23):  # 1935-1939: not yet listed
        ws.cell(r, col).value = vals(NOTYET)[(col - 3) % 5]
    state = NOTYET
    for y, col in YEARS.items():
        state = changes.get(y, DEAD if state == DEAD else state)
        write(r, col, vals(state))
    r += 1

wb.save(DST)
print('saved', DST, 'new rows', last + 1, '-', r - 1)
