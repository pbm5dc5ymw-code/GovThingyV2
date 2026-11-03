"""Agency-not-position review of 1944-1950, applied after format_review_1944_1948.py; re-run it if any earlier
script is re-run.

The user's three requirements for a row-year:
  1. the name is an agency, not just a position;
  2. the formatting sets the agency apart (usually a bold heading; in these editions also a run-in in capitals
     and small capitals);
  3. a description of the agency follows.
So "OFFICE OF THE CHIEF OF ENGINEERS.--The Chief of Engineers is ..." passes (the set-off name is an office),
but "THE INSPECTOR GENERAL.--" or "CHIEF OF ENGINEERS.--" fails (the set-off name is an official). A unit only in
an "as follows" list ("OFFICES: Office of the Attorney General ...") or an officials list fails as before.
Kept: the Navy's bold "The Chief of Naval Operations" (1946-48), because the text under it describes the Office
("The Office of the Chief of Naval Operations is organized ...").

Page numbers are printed pages (PDF page = printed + 8; 1946 and 1949 = printed + 10).
"""
import openpyxl

PATH = 'data/Spreadsheet 1935-1950.xlsx'
DEAD = 'DEAD'
JUS = 'Department of Justice'

# Row -> years where the only set-off text names a position (or nothing is set off).
REMOVE = {
    95: range(1944, 1950),      # "THE ATTORNEY GENERAL.--" (1948 p. 202); the office only in the OFFICES list
    639: range(1944, 1950),     # "THE SOLICITOR GENERAL.--" (1949 p. 193)
    87: range(1944, 1950),      # "THE ASSISTANT TO THE ATTORNEY GENERAL.--" (1949 p. 194)
    88: range(1944, 1950),      # "THE ASSISTANT SOLICITOR GENERAL.--" (1949 p. 196)
    693: range(1946, 1950), 694: range(1946, 1950),   # "UNITED STATES ATTORNEYS AND MARSHALS": officials and addresses only
    110: range(1944, 1951),     # "LEGAL ADVISER.--" every year (1950 p. 72)
    214: (1948,),               # Navy Under Secretary: officials list only (p. 164)
    215: (1948, 1949, 1950),    # Navy Office of the Secretary: officials list; bold "Secretary of the Navy" (1950 p. 137)
    60: (1949,), 698: (1949,),  # Commerce and Interior Offices of the Secretary: a sentence and a unit list (pp. 288, 222)
    680: (1946,),               # bold "Director of Liquidation" (p. 63): a position
    275: (1944, 1945, 1946, 1950),   # "THE INSPECTOR GENERAL.--" (1946 p. 199), "INSPECTOR GENERAL.--" (1950 p. 129)
    283: range(1946, 1951),     # bold "Director of Information" (1946 p. 197), "Chief of Information" (1947 p. 160), run-ins 1949-50
    536: (1949, 1950),          # "PROVOST MARSHAL GENERAL.--" (1949 p. 158, 1950 p. 130)
    # 1950 Army: every chief is a run-in naming the official (pp. 129-131)
    274: (1950,), 276: (1950,), 277: (1950,), 278: (1950,), 279: (1950,), 280: (1950,), 291: (1950,),
    292: (1950,), 293: (1950,), 294: (1950,), 296: (1950,), 539: (1950,), 638: (1950,), 775: (1950,), 777: (1950,),
}

# Row -> {year: parent}
RESTORE = {
    87: {1950: JUS},            # bold "Office of the Deputy Attorney General" with a description (1950 p. 158); renamed
}

wb = openpyxl.load_workbook(PATH)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}


def put(r, y, c):
    v = ['Dead'] * 5 if c == DEAD else ['No', 'Yes', c, 'None', 'No']
    for i, x in enumerate(v):
        ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = x


for r, ys in REMOVE.items():
    for y in ys:
        put(r, y, DEAD)
for r, plan in RESTORE.items():
    for y, c in plan.items():
        put(r, y, c)

# Dead rows are hidden; a row live in 1950 is shown.
for r in set(REMOVE) | set(RESTORE):
    ws.row_dimensions[r].hidden = str(ws.cell(r, hdr['1950_cabinet_agency']).value).strip() == 'Dead'

# No live row may name a removed row as its parent.
names = {str(ws.cell(r, 1).value).strip() for r in REMOVE}
for y in range(1944, 1951):
    c = hdr[f'{y}_cabinet_agency']
    for r in range(3, ws.max_row + 1):
        p = str(ws.cell(r, c + 3).value or '').strip()
        if p in names:
            live = [q for q in range(3, ws.max_row + 1) if str(ws.cell(q, 1).value).strip() == p
                    and str(ws.cell(q, c).value).strip() != 'Dead']
            assert live, (y, r, p)

wb.save(PATH)
print('row-years removed', sum(len(v) for v in REMOVE.values()), 'restored', sum(len(v) for v in RESTORE.values()))
