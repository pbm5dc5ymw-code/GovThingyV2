"""Format-priority review of 1944-1948, applied to the latest workbook after fill1950.py; re-run it if any
fill script is re-run. Rows are the same as fill1950.py produces.

Same rule as 1949-50: a unit needs its name set off and then a description of the unit. These editions are
one column, so "set off" reads differently from the two-column 1950 edition:
  - a bold centered heading (large with address and officials for agencies, smaller inside a section), or
  - a run-in at the start of a paragraph in capitals and small capitals: "Division of Traffic.--This Division ...".
  - The description may be only the territory a division covers ("Division of Near Eastern Affairs.--Egypt, ..."):
    in these editions that is how every State geographic division is described, so it counts.
  - Several units named together in one run-in count ("Policy and Post-War Programs Committees.--").
Not set off: a name in a plain paragraph, in the "as follows" unit lists set in two columns, in an officials
list, or on a chart. Charts begin these editions' back matter and still only decide parents.

Page numbers are printed pages (PDF page = printed + 8; 1946 = printed + 10). Each list below was checked
against a scan of every heading and run-in in the edition; the uncertain cases were read on the page image.
"""
import openpyxl

PATH = 'data/Spreadsheet 1935-1950.xlsx'
ALL = (1944, 1945, 1946, 1947, 1948)
AG, COM, JUS, LAB, ST = ('Department of Agriculture', 'Department of Commerce', 'Department of Justice',
                         'Department of Labor', 'Department of State')
INT, NAVY, TRE, WAR, PO = ('Department of the Interior', 'Department of the Navy', 'Department of the Treasury',
                           'Department of War', 'Post Office Department')
NME = 'National Military Establishment'
CABINET = {AG, COM, JUS, LAB, ST, INT, NAVY, TRE, WAR, PO, NME}
SS, RFC, CSC = 'Special Staff', 'Reconstruction Finance Corporation', 'United States Civil Service Commission'
WMC, NA, USES = 'War Manpower Commission', 'National Archives', 'United States Employment Service'
DEAD = 'DEAD'

# Row -> years with no set-off text. Reason in the comment.
REMOVE = {
    4: (1944,),                        # AAA: renamed Agricultural Adjustment Agency; sentence only (p. 364)
    60: ALL, 215: (1944, 1945, 1946, 1947), 698: (1946, 1947, 1948),   # Offices of the Secretary (Commerce, Navy, Interior): lists only
    95: ALL,                           # Office of the Attorney General: sentence under "The Attorney General.--" (1948 p. 202)
    207: (1946, 1947, 1948),           # Puerto Rico Reconstruction Administration: officials and lists only (1947 p. 232)
    214: (1944, 1945, 1946, 1947), 217: ALL, 228: ALL,                 # Navy Under and Assistant Secretaries: officials only
    281: (1944, 1945), 537: (1944,),   # Army Morale Branch, Armored Force: not printed as units
    283: (1947, 1948),                 # Bureau of Public Relations: gone; "Chief of Information.--" names an official
    311: (1944, 1945),                 # Federal Committee on Apprenticeship: appendix only
    319: (1944,),                      # Federal Farm Mortgage Corporations: no such name printed
    390: ALL,                          # Division of Post Office Service: no such name printed
    391: (1946, 1947, 1948),           # Division of Postmasters: plain sentence (1948 p. 220)
    408: (1945, 1946, 1947, 1948), 643: (1945, 1946, 1947, 1948),      # Cost Ascertainment, Methods and Procedures: lists only (1948 p. 218)
    486: (1945, 1946), 487: (1945, 1946),   # Social Security Board service bureaus: plain paragraphs (1946 p. 393)
    545: (1944, 1945),                 # Training Within Industry: sentence in the War Manpower section (1945 p. 94)
    546: (1944,),                      # War Public Works Program: program paragraph only
    551: (1944,), 552: (1944,),        # Food Distribution, Food Production Administrations: history sentences in the Office of Distribution section
    553: (1944,),                      # Budget and Administrative Management: list and sentence (pp. 317-318); heading from 1945
    560: (1944, 1945, 1946),           # Office of Community War Services: plain paragraph (1944 p. 436, 1946 p. 382)
    614: (1946, 1947, 1948),           # Commodities Division: not printed in State; a Commerce list only
    656: (1945,), 657: (1945,),        # International Information offices: list and transfer sentences
    658: (1945,), 659: (1945, 1946),   # Staff Committee, Coordinating Committee: officials and a sentence
    682: (1948,),                      # National Intelligence Authority: "ceased to exist" (p. 62)
    689: (1946,), 692: (1946, 1947, 1948),  # Labor Information, Treasury Public Relations: chart or not printed
    695: (1946,), 696: (1946,), 697: (1946,),   # War Secretariat, Strength Accounting, Personnel Branch: chart only (p. 578)
    699: (1946, 1947, 1948), 700: (1946, 1947), 701: (1946, 1947),     # Interior staff offices: lists and chart only
    708: (1946,),                      # Office of War Property Distribution: officials and chart only
    709: (1946, 1947, 1948), 710: (1946, 1947, 1948), 711: (1946,), 712: (1946, 1947, 1948),   # Federal Works Agency staff offices: chart only
    740: (1948,),                      # Map Intelligence: run-in in 1947 only
    766: (1947, 1948), 767: (1947, 1948), 768: (1947, 1948), 769: (1947, 1948), 770: (1947, 1948), 771: (1947, 1948),   # Foreign Liquidation divisions: named together in a sentence (1948 p. 92)
    773: (1947, 1948),                 # Post Office Administrative Services: list only (1948 p. 217)
    784: (1947, 1948),                 # Interior Geography: chart only
    792: (1947,), 794: (1948,),        # Appeals Council, Office of Special Services: chart only
    797: (1947,), 799: (1947,), 800: (1947, 1948), 801: (1947,), 802: (1947,),   # FSA staff offices: officials and chart only; Administration has a heading from 1948
    824: (1948,),                      # Coordinator for Foreign Aid: not printed
    840: (1948,), 841: (1948,),        # Interior Programs, Research and Marketing Administrator: chart and officials only
    845: (1948,), 846: (1948,), 847: (1948,),   # PHS bureaus: "Research.--" describes a program; no headings in 1948 (p. 384)
}

# Row or name -> {year: parent} for rows set off in print but coded dead or not listed.
RESTORE = {
    173: {1944: INT, 1945: INT},       # Alaska Road Commission run-in (1944 p. 332)
    197: {1946: INT, 1947: INT, 1948: INT},   # Office of the Chief Clerk heading (1946 p. 275)
    270: {1944: WAR, 1945: WAR},       # Special Staff heading (1944 p. 252)
    775: {1944: SS, 1945: SS, 1946: SS}, 777: {1946: SS}, 778: {1944: SS, 1946: SS},
    779: {1945: SS}, 780: {1945: SS, 1946: SS},   # Special Staff divisions: run-ins (1944 p. 252, 1946 pp. 199-200)
    811: {1944: WAR, 1945: WAR, 1946: WAR, 1947: WAR}, 813: {1947: WAR},   # Aeronautical Board, Explosives Safety Board (1947 pp. 170-172)
    854: {1948: NME},                  # Military Liaison Committee heading (p. 161)
    112: {y: CSC for y in ALL},        # Board of Appeals and Review run-in (1948 p. 494)
    433: {1945: RFC}, 434: {1945: RFC}, 435: {1945: RFC}, 436: {1945: RFC},   # RFC subsidiaries: run-ins (pp. 407-408)
    44: {1948: RFC}, 566: {1948: RFC},  # War Damage Corporation, U. S. Commercial Company: run-ins (p. 465)
    686: {1947: LAB},                  # Shipbuilding Stabilization Committee heading (p. 332)
    'National Archives Trust Fund Board': {y: NA for y in ALL},   # run-in (1948 p. 443)
    'Institute of Inter-American Affairs': {1947: ST, 1948: ST},  # heading (1948 p. 100)
    'Division of Liquidation': {1948: COM}, 'Office of Industry Cooperation': {1948: COM},   # headings (pp. 292-293)
    'Office of Program Planning (Commerce)': {1946: COM, 1947: COM, 1948: COM},   # heading (1946 p. 319)
    'Office of Publications (Commerce)': {1947: COM, 1948: COM},  # heading (1947 p. 296)
}

# New rows at the bottom: set-off units with no row, at the depth 1935-48 already use.
NEW = [
    ('Bureau of Manpower Utilization', {1944: WMC, 1945: WMC}),   # 1944 p. 103
    ('Bureau of Placement', {1944: WMC, 1945: WMC}),              # 1944 p. 104
    ('Bureau of Training', {1944: WMC, 1945: WMC}),               # 1944 p. 105
    ('Veterans Employment Service', {1945: WMC, 1946: USES, 1947: USES, 1948: LAB}),   # 1945 p. 95, 1946 p. 359, 1947 p. 336, 1948 p. 327
    ('Joint Contract Termination Board', {1944: 'Office of War Mobilization'}),   # p. 98
    ('American Mexican Claims Commission', {1944: ST, 1946: ST}),  # 1944 p. 206
    ('Operations Board (Post Office)', {1944: PO, 1945: PO}),     # 1944 p. 273
    ('Division of Budget and Accounts (Commerce)', {1944: COM}),  # p. 388
    ('Division of Purchases and Sales (Commerce)', {1944: COM}),  # p. 388
    ('Office of Declassification and Technical Services', {1946: COM}),   # p. 320
    ('Joint Research and Development Board', {1947: WAR}),        # p. 173
]


def vals(c):
    if c == DEAD: return ['Dead'] * 5
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


wb = openpyxl.load_workbook(PATH)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
years = sorted(int(h[:4]) for h in hdr if h.endswith('_cabinet_agency'))
last = max(r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value)
byname = {str(ws.cell(r, 1).value).strip(): r for r in range(last, 2, -1) if ws.cell(r, 1).value}


def put(r, y, c):
    for i, v in enumerate(vals(c)):
        ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v


def code(r, y):
    v = [str(ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value or '').strip() for i in range(5)]
    if v[0] == 'Dead': return DEAD
    if v[1] == 'Yes': return v[2] if v[2] != 'None' else v[3]
    return 'CAB' if v[0] == 'Yes' else 'OTHER' if v[4] == 'Yes' else None


for r, ys in REMOVE.items():
    for y in ys:
        put(r, y, DEAD)
for key, plan in RESTORE.items():
    r = byname[key] if isinstance(key, str) else key
    for y, c in plan.items():
        put(r, y, c)

r = last
for name, plan in NEW:
    r += 1
    ws.cell(r, 1).value = name
    ws.cell(r, 2).value = 'Executive'
    end = max(plan)
    for y in years:
        if y in plan:
            put(r, y, plan[y])
        elif y > end:
            put(r, y, DEAD)
        else:
            for i, v in enumerate(['No', 'No', 'None', 'None', 'No']):
                ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v
    byname[name] = r
    ws.row_dimensions[r].hidden = True          # ended before 1950

# A unit whose parent came off moves up to the parent's own parent. Names repeat across departments
# ("Bureau of Accounts"), so a parent counts as gone only if no row of that name is live that year.
UP = {'Office of International Information and Cultural Affairs': ST}   # no earlier year to take its parent from
allrows = {}
for r in range(3, ws.max_row + 1):
    if ws.cell(r, 1).value:
        allrows.setdefault(str(ws.cell(r, 1).value).strip(), []).append(r)


def gone(name, y):
    return name in allrows and all(code(r, y) in (DEAD, None) for r in allrows[name])


moved = 0
for y in ALL:
    for r in range(3, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        p = code(r, y)
        if not p or p in ('CAB', 'OTHER', DEAD) or not gone(p, y):
            continue
        up, hops = p, 0
        while gone(up, y) and hops < 5:
            hops += 1
            if up in UP:
                up = UP[up]
                continue
            found = None
            for e in range(y - 1, 1934, -1):
                if f'{e}_cabinet_agency' in hdr:
                    found = next((code(q, e) for q in allrows[up] if code(q, e) not in (DEAD, None)), None)
                    if found:
                        break
            assert found, (y, r, up)
            up = found
        if up == 'OTHER':
            for i, v in enumerate(['No', 'No', 'None', 'None', 'Yes']):
                ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v
        else:
            put(r, y, up)
        moved += 1
        print('moved', y, r, ws.cell(r, 1).value, ':', p, '->', up)

wb.save(PATH)
print('removed', sum(len(v) for v in REMOVE.values()), 'restored', sum(len(v) for v in RESTORE.values()),
      'new rows', len(NEW), 'moved', moved)
