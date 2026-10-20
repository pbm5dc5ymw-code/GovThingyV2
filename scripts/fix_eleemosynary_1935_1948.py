"""Fill in the four former Interior "eleemosynary institutions" for 1935-1948.

Freedmen's Hospital, Saint Elizabeths Hospital, Howard University and the Columbia Institution for the
Deaf are printed in every Manual but were missing (no row before 1946/1947). Each year was checked on
the page image (PDF page numbers in the comments). Applied to the latest workbook after the fill scripts;
re-run it if any fill script is re-run.

  1935-1939  Interior unit list, indented under "Eleemosynary institutions:" (a grouping label, not a
             unit) -> child of Interior.  1935 p.111 (printed 79), 1939 p.152 (printed 140).
  1940-1941  Moved to the Federal Security Agency by Reorganization Plan IV (June 30, 1940); separate
             solid-line boxes on the FSA chart -> child of FSA.  1940 chart p.619, 1941 chart p.626.
  1942       FSA chart: Saint Elizabeths and Freedmen's printed inside the Public Health Service box ->
             child of PHS; Howard and Columbia separate boxes -> FSA.  1942 chart p.589.
  1943       No chart. FSA text: "the Public Health Service, under which are Freedmen's and St.
             Elizabeths"; both sections approved by the Surgeon General -> PHS.  Howard and Columbia have
             their own headings in the FSA section -> FSA.  1943 pp.436, 440.
  1944-1945  "Public Health Service, under which is Freedmen's Hospital; St. Elizabeths Hospital; ..."
             Freedmen's approved by the Surgeon General -> PHS; St. Elizabeths approved by its own
             Superintendent, listed separately -> FSA.  Howard, Columbia headings in FSA section -> FSA.
  1946       Same text for the hospitals (Freedmen's -> PHS, St. Elizabeths -> FSA, chart p.595 solid).
             Chart joins Howard and Columbia by DASHED lines ("Semiofficial agencies") -> standalone.
  1947       Freedmen's section approved by the Surgeon General -> PHS (was FSA). Others unchanged; Howard
             and Columbia drawn inside the FSA Education box -> FSA.
  1948       Howard and Columbia in the FSA "Federally Aided Corporations" box -> FSA.
"""
import openpyxl

PATH = 'data/Spreadsheet 1935-1948.xlsx'
INT, FSA, PHS = 'Department of the Interior', 'Federal Security Agency', 'Public Health Service'
OTHER = 'OTHER'
CABINET = {INT}


def vals(c):
    if c == OTHER: return ['No', 'No', 'None', 'None', 'Yes']
    if c in CABINET: return ['No', 'Yes', c, 'None', 'No']
    return ['No', 'Yes', 'None', c, 'No']


EARLY = {y: INT for y in (1935, 1936, 1937, 1939)}
PLAN = {
    "Freedmen's Hospital": {**EARLY, 1940: FSA, 1941: FSA, 1942: PHS, 1943: PHS, 1944: PHS, 1945: PHS, 1946: PHS, 1947: PHS},
    'St. Elizabeths Hospital': {**EARLY, 1940: FSA, 1941: FSA, 1942: PHS, 1943: PHS, 1944: FSA, 1945: FSA},
    'Howard University': {**EARLY, 1940: FSA, 1941: FSA, 1942: FSA, 1943: FSA, 1944: FSA, 1945: FSA, 1946: OTHER, 1947: FSA, 1948: FSA},
    'Columbia Institution for the Deaf': {**EARLY, 1940: FSA, 1941: FSA, 1942: FSA, 1943: FSA, 1944: FSA, 1945: FSA, 1946: OTHER, 1947: FSA, 1948: FSA},
}

wb = openpyxl.load_workbook(PATH)
ws = wb.worksheets[0]
hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
years = sorted(int(h[:4]) for h in hdr if h and h.endswith('_cabinet_agency') and ws.cell(3, hdr[h]).value is not None)
rows = {str(ws.cell(r, 1).value).strip(): r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value}
last = max(rows.values())
for name, plan in PLAN.items():
    r = rows.get(name)
    if r is None:
        last += 1
        r = last
        ws.cell(r, 1).value = name
        ws.cell(r, 2).value = 'Executive'
        for y in years:
            for i, v in enumerate(['No', 'No', 'None', 'None', 'No']):
                ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v
        print('added row', r, name)
    for y, code in plan.items():
        for i, v in enumerate(vals(code)):
            ws.cell(r, hdr[f'{y}_cabinet_agency'] + i).value = v
    print(name, r, {y: plan[y] for y in sorted(plan)})
wb.save(PATH)
