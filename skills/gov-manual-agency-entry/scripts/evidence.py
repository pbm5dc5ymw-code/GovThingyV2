"""Evidence table: how each sheet row's name is printed in one manual edition.

Usage: python evidence.py <manual.txt> <sheet.xlsx> <year> [<prev_year>] [<out.tsv>]
For every named row (live, blank or dead) it searches the body of the manual (before Appendix A) and
classifies each place the name appears, strongest first:
  HEADING     the name alone on a line, followed by an address/phone/OFFICIALS/MEMBERS line
  SUBHEAD     the name alone on an all-capitals line (centered subheading inside another section)
  LIST        the name alone (or as a short line) on a page with an "as follows"/"includes" unit list
  OFFICIAL    "Chief/Director/... , <name>" line in an officials list - a job title only; NOT enough
  DESCRIBED   the name followed by a description of the unit itself: "<NAME>.—This office ..." or
              "<Name> The <Name> is/was ..." (a heading without an address)
  RUNIN       "<NAME>.—" run-in paragraph heading that describes a program or a person, not the unit
  SENTENCE    only inside running text
  NONE        not found (check spelling; the row name may be misspelled)
The text locates evidence; confirm SENTENCE/NONE live rows and HEADING/LIST blank or dead rows on the page
image before changing them. Prints a summary of the rows that need a decision; writes all rows to TSV.
"""
import difflib, re, sys
import openpyxl

ADDR = re.compile(r'(Building|Street|Avenue|NW\.|SW\.|REpublic|EXecutive|NAtional|DIstrict|DUpont|OFFICIALS|MEMBERS|'
                  r'Telephone|Washington|Room \d|Branch \d|Officials|Members|BOARD OF|DIRECTORS|COMMISSIONERS)')
OFFL = re.compile(r'^(Chief|Director|Commissioner|Administrator|Chairman|Head|Manager|Superintendent|President|'
                  r'Executive Director|Governor|Deputy|Assistant|Secretary|General Manager|Coordinator|Librarian)\b')
RANK = ['HEADING', 'SUBHEAD', 'DESCRIBED', 'LIST', 'OFFICIAL', 'RUNIN', 'SENTENCE', 'NONE']
UNITWORD = r'(office|division|bureau|staff|board|service|section|unit|committee|council|commission|agency|administration|corporation|branch|program)'


def norm(s):
    s = str(s).lower().replace('’', "'")
    s = re.sub(r'\([^)]*\)', ' ', s)          # drop "(Interior)" style suffixes
    s = re.sub(r'[^a-z ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def load(txt):
    t = open(txt, encoding='utf-8').read()
    parts = re.split(r'=====PAGE (\d+)\n', t)[1:]
    return {int(parts[i]): parts[i + 1] for i in range(0, len(parts), 2)}


def offset(P):
    import collections
    c = collections.Counter()
    for k, v in P.items():
        L = [l.strip() for l in v.split('\n') if l.strip()]
        for l in L[:2] + L[-2:]:
            m = re.match(r'^(\d{1,3})\b', l) or re.search(r'\b(\d{1,3})$', l)
            if m and 5 < int(m.group(1)) < 700:
                c[k - int(m.group(1))] += 1
    return c.most_common(1)[0][0]


def body_pages(P):
    app = [k for k, v in P.items() if re.search(r'^\s*APPENDIX A\s*$', v, re.M) and k > 300]
    end = min(app) if app else max(P)
    # organization charts are drawings, not printed evidence: drop everything from the chart list onward
    charts = [k for k, v in P.items() if 'LIST OF ORGANIZATION CHARTS' in v and k > 300]
    if charts:
        end = min(end, min(charts))
    return {k: v for k, v in P.items() if k < end}


def classify(lines, i, key, listpage):
    l = lines[i].strip()
    nl = norm(l)
    nxt = ' '.join(x.strip() for x in lines[i + 1:i + 4])
    letters = [c for c in l if c.isalpha()]
    upper = letters and sum(c.isupper() for c in letters) / len(letters) > 0.9
    if nl in (key, 'the ' + key) or nl.endswith(' ' + key) and len(nl) - len(key) < 6:
        if ADDR.search(nxt) and not upper:
            return 'HEADING'
        if upper:
            return 'SUBHEAD'
        if listpage:
            return 'LIST'
    if OFFL.match(l) and key in nl:
        return 'OFFICIAL'
    m = re.match(r'^(.{5,120}?)\s*\.\s*[—-]', l)
    if m and norm(m.group(1)) in (key, 'the ' + key):
        return 'RUNIN'
    if listpage and key in nl and len(nl) <= len(key) + 45 and not re.search(r'[.;]\s', l):
        return 'LIST'
    return 'SENTENCE'


def main():
    txt, xlsx, year = sys.argv[1], sys.argv[2], int(sys.argv[3])
    prev = int(sys.argv[4]) if len(sys.argv) > 4 else None
    out = sys.argv[5] if len(sys.argv) > 5 else f'evidence_{year}.tsv'
    P = load(txt)
    off = offset(P)
    B = body_pages(P)
    pages = {k: v.split('\n') for k, v in B.items()}
    flat = {k: norm(re.sub(r'-\n', '', v)) for k, v in B.items()}
    listp = {k for k, v in B.items() if re.search(r'as follows|includes the|following (offices|bureaus|units|divisions|constituent)|comprises', v)}
    vocab = set(w for v in flat.values() for w in v.split())
    vlist = sorted(vocab)
    # unit-list blocks: text after "as follows"/"includes the" up to the next ACTIVITIES heading or 2500 chars
    blocks = {}
    for k, v in B.items():
        t = re.sub(r'-\n', '', v)
        for m in re.finditer(r'(as follows|includes the|comprises the|following offices|following (constituent )?units)', t):
            seg = t[m.end():m.end() + 2500]
            seg = re.split(r'\n\s*ACTIVITIES\b|CREATION AND', seg)[0]
            blocks.setdefault(k, []).append(norm(seg))
        if k + 1 in B and k in blocks:    # lists that continue onto the next page
            blocks[k].append(norm(B[k + 1][:1500]))
    TITLE = r'(chief|director|commissioner|administrator|chairman|head|manager|superintendent|president|governor|coordinator|executive|librarian)'

    ws = openpyxl.load_workbook(xlsx).worksheets[0]
    hdr = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(2, c).value}
    pc = hdr.get(f'{prev}_cabinet_agency') if prev else None
    rows = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name or not str(name).strip():
            continue
        kind = str(ws.cell(r, 2).value or '').strip()
        if kind in ('Legislative', 'Judicial'):
            continue
        key = norm(name)
        status = ''
        if pc:
            v = [str(ws.cell(r, pc + i).value or '').strip() for i in range(5)]
            status = ('blank' if not any(v) else 'Dead' if v[0] == 'Dead' else 'CAB' if v[0] == 'Yes' else
                      (v[2] if v[2] not in ('None', '') else v[3]) if v[1] == 'Yes' else 'OTHER' if v[4] == 'Yes' else 'notlisted')
        used = key
        hits = [k for k in flat if key and key in flat[k]]
        if not hits and len(key) > 10:      # misspelled row names: correct each word against the manual's vocabulary
            fixed = []
            for w in key.split():
                if w in vocab or len(w) < 4:
                    fixed.append(w)
                else:
                    c = difflib.get_close_matches(w, vlist, n=1, cutoff=0.8)
                    fixed.append(c[0] if c else w)
            fixed = ' '.join(fixed)
            if fixed != key:
                used = fixed
                hits = [k for k in flat if used in flat[k]]
        core = re.sub(r'^(the )?(office|bureau) of (the )?', '', used)
        core = re.sub(r'^(the )?(office|bureau|division) of (the )?', '', core)
        best, where = 'NONE', {}
        for k in hits:
            L = pages[k]
            found = False
            for i in range(len(L)):
                cand = [norm(L[i])]
                if i + 1 < len(L):
                    cand.append(norm(L[i] + ' ' + L[i + 1]))
                if any(used in c for c in cand):
                    found = True
                    j = i
                    c = classify(L, i, used, k in listp)
                    if c == 'SENTENCE' and i + 1 < len(L) and used not in norm(L[i]):
                        c2 = classify([L[i] + ' ' + L[i + 1]] + L[i + 2:], 0, used, k in listp)
                        c = c2
                    where.setdefault(c, []).append(k - off)
            if not found:
                where.setdefault('SENTENCE', []).append(k - off)
            if re.search(re.escape(used) + r' ?\.? ?(—|-)? ?this ' + UNITWORD, flat[k]) or \
               re.search(re.escape(used) + r' the ' + re.escape(used) + r' (is|was|has|serves|performs|carries|discharges)', flat[k]):
                where.setdefault('DESCRIBED', []).append(k - off)
            if any(used in b for b in blocks.get(k, [])):
                where.setdefault('LIST', []).append(k - off)
            if re.search(TITLE + r'(\s+\w+){0,2}\s+(of\s+)?(the\s+)?' + re.escape(used), flat[k]):
                where.setdefault('OFFICIAL', []).append(k - off)
        if core != used and len(core) > 8:   # "Office of the Chief of Finance" is printed as the official "Chief of Finance"
            for k in flat:
                if re.search(r'(^| )(the )?' + re.escape(core) + r'( |$)', flat[k]) and re.search(re.escape(core) + r' [a-z]+ [a-z]', flat[k]):
                    if k in listp or 'officials' in flat[k]:
                        where.setdefault('OFFICIAL', []).append(k - off)
        for c in RANK:
            if c in where:
                best = c
                break
        pg = sorted(set(where.get(best, [])))[:6]
        rows.append((r, str(name).strip(), status, best, pg, '' if used == key else used,
                     ', '.join(f'{c}:{len(set(v))}' for c, v in where.items())))

    with open(out, 'w') as f:
        f.write('row\tname\tprev_status\tevidence\tprinted_pages\tmatched_as\tall\n')
        for x in rows:
            f.write('\t'.join(map(str, x)) + '\n')
    live = lambda s: s not in ('', 'blank', 'Dead', 'notlisted')
    weak = [x for x in rows if live(x[2]) and x[3] in ('SENTENCE', 'NONE', 'RUNIN', 'OFFICIAL')]   # a job title is not a unit
    strong_off = [x for x in rows if not live(x[2]) and x[3] in ('HEADING', 'SUBHEAD', 'LIST')]
    print(f'{len(rows)} rows; PDF page = printed + {off}; written to {out}')
    print(f'\nLive last year but weak/no evidence this year ({len(weak)}) — confirm on page or mark Dead:')
    for x in weak:
        print(f'  {x[0]:>4} {x[1][:60]:60} [{x[2][:35]}] {x[3]} pp{x[4]} {x[5]}')
    print(f'\nNot live last year but printed as a unit this year ({len(strong_off)}) — check and fill:')
    for x in strong_off:
        print(f'  {x[0]:>4} {x[1][:60]:60} [{x[2]}] {x[3]} pp{x[4]}')


if __name__ == '__main__':
    main()
