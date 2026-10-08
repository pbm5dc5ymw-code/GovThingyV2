"""Compare two agency spreadsheets (e.g. machine-entered vs. hand-entered) year by year.

Usage: python compare_sheets.py <mine.xlsx> <hand.xlsx> <hand_tab_name> <year> [<year> ...]
Rows are matched by agency name (letters only, case-insensitive; nearest row number breaks ties).
Prints agreement per year three ways and writes every difference to compare_differences.tsv.
  exact       - all matched rows identical in category and parent
  dead~blank  - same, but "Dead" vs. blank (dead rows left empty) counts as agreement
  live rows   - only rows where either sheet shows a live agency (the strictest, most useful number)
Then classify each difference before calling it an error: reading error, convention, edition, or not entered.
"""
import collections, re, sys
import openpyxl


def load(path, tab=None):
    wb = openpyxl.load_workbook(path)
    ws = wb[tab] if tab else wb.worksheets[0]
    hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
    out = []
    for r in range(3, ws.max_row + 1):
        n = ws.cell(r, 1).value
        if not n or not str(n).strip():
            continue
        g = lambda c: '' if ws.cell(r, c).value is None else str(ws.cell(r, c).value).strip()
        out.append(dict(r=r, n=re.sub(r'\s+', ' ', str(n)).strip(), ws=ws, hdr=hdr, g=g))
    return out


def cat(x, y):
    c = x['hdr'].get(f'{y}_cabinet_agency')
    if not c: return 'blank'
    t = tuple(x['g'](c + i).replace('Department of Interior', 'Department of the Interior')
              .replace('Department of Treasury', 'Department of the Treasury').replace('Department of War', 'War Department') for i in range(5))
    if not any(t): return 'blank'
    if t[0] == 'Dead': return 'Dead'
    if t[0] == 'Yes': return 'CAB'
    if t[1] == 'Yes': return 'CH(' + (t[2] if t[2] not in ('None', '') else t[3]) + ')'
    if t[4] == 'Yes': return 'OTHER'
    return 'notyet'


mine, hand = load(sys.argv[1]), load(sys.argv[2], sys.argv[3])
years = [int(y) for y in sys.argv[4:]]
key = lambda n: re.sub(r'[^a-z]', '', n.lower())
idx = collections.defaultdict(list)
for x in hand:
    idx[key(x['n'])].append(x)
used, pairs = set(), []
for m in mine:
    c = [x for x in idx.get(key(m['n']), []) if id(x) not in used]
    if c:
        b = min(c, key=lambda x: abs(x['r'] - m['r']))
        used.add(id(b))
        pairs.append((m, b))
print('matched', len(pairs), '| only in mine', len(mine) - len(pairs), '| only in hand', len(hand) - len(pairs))
with open('compare_differences.tsv', 'w') as f:
    f.write('year\trow\tagency\tmine\thand\n')
    for y in years:
        raw = adj = live = ok = 0
        for m, h in pairs:
            a, b = cat(m, y), cat(h, y)
            raw += a == b
            adj += a == b or (a == 'Dead' and b in ('blank', 'Dead'))
            if a not in ('blank', 'Dead', 'notyet') or b not in ('blank', 'Dead', 'notyet'):
                live += 1
                ok += a == b
            if a != b and not (a == 'Dead' and b == 'blank'):
                f.write(f'{y}\t{m["r"]}\t{m["n"]}\t{a}\t{b}\n')
        n = len(pairs)
        print(f'{y}: exact {raw / n:.0%} | dead~blank {adj / n:.0%} | live rows {ok}/{live} = {ok / max(live, 1):.0%}')
print('differences written to compare_differences.tsv')
