---
name: gov-manual-agency-entry
description: Enter one year of U.S. federal executive agencies from a United States Government Manual (GPO GOVMAN PDF) into the agency-by-year spreadsheet — coding each agency as cabinet, child (with cabinet or other parent), other, not listed, or Dead — and check it against hand-entered data. Use this whenever the user asks to "do 1944", "fill in the next year", add a Government Manual year, recheck or correct agency years, compare the spreadsheet with hand-entered or feedback data, list reading errors, or anything involving GOVMAN PDFs, agency hierarchies, parent agencies, or the GovThingy agency spreadsheet — even if they don't say "skill".
---

# Government Manual → agency-by-year spreadsheet

The project tracks every executive-branch agency year by year (1935 onward) from the *United States
Government Manual*, recording whether each agency is a cabinet department, a child of a cabinet department,
a child of another agency, a standalone ("other") agency, not yet listed, or Dead. Accuracy is the whole
point: the dataset is compared against careful hand entry, and **reading errors matter most**.

Most past errors came from process (wrong edition, rules guessed from one example, carrying last year
forward, reading extracted text instead of the page layout) — not from misreading words. The workflow
below is ordered to prevent exactly those failures. Read `references/coding-rules.md` before coding a
year, and skim `references/lessons.md` once.

## Tools
- Python with `pypdf`, `pymupdf`, `openpyxl` (create a venv if `pip` is blocked:
  `python3 -m venv venv && ./venv/bin/pip install pypdf pymupdf openpyxl`).
- `scripts/manual_text.py` — extract text, contents/offset, Appendix A changes, section headers, unit
  lists, term search, name presence check. Text is for **finding** things only.
- `scripts/render_pages.py` — render PDF pages to PNG; read them with the Read tool. Structure (who is
  whose parent) is decided **only** from these images.
- `scripts/fill_year_template.py` — copy to `scripts/fill<YEAR>.py`; it carries the previous year
  forward and applies your page-cited changes, appends new rows, and hides dead rows.
- `scripts/check_sheet.py` — parents live, duplicates agree, earlier years untouched, hidden rows.
- `scripts/compare_sheets.py` — agreement with a hand-entered tab, differences to TSV.

## Workflow for one year

### 1. Settle the rules and the source
1. Open the workbook and read **every tab** — especially *Instructions*, *FAQs*, *Sources*, *Questions*.
   Restate any rule that differs from `references/coding-rules.md` and follow the workbook.
2. Find every edition for the year on govinfo (API URL in coding-rules §7). Use the **latest edition
   dated in that year**. If it isn't downloaded, download it (check `content-length`; large files may need
   `curl -C -` resumes until the size matches) or ask the user for it. Say which edition you used.
3. Note any OPEN conventions (coding-rules §8) you will need; use the defaults and flag them.

### 2. Map the manual
1. `manual_text.py extract <pdf> <txt>`, then `contents <txt>` (gives the PDF→printed page offset and
   whether organization charts exist in this edition).
2. `heads <txt>` to map sections (Executive Office, emergency/war agencies, departments, independents,
   quasi-official, charts, appendix).
3. `appendix <txt> <year>` — every agency abolished, transferred, consolidated or renamed recently.
   Render the relevant appendix pages if the two-column text is garbled.

### 3. Read the structure from page images
For each department and umbrella agency (Executive Office, OEM, Federal Security/Works/Loan Agencies,
National Housing Agency, etc.):
1. Locate pages: `lists <txt>` for "as follows:" unit lists; `find <txt> "<name>"` for sections; the
   chart list near the back (if charts exist).
2. `render_pages.py <pdf> <outdir> 115 <pages...>` and **Read each PNG**. Decide parents from indentation,
   colons, boxes, solid vs. dashed lines and footnotes (coding-rules §4).
3. Rebuild that department's unit list **from this year's page** — do not start from last year. Compare
   with last year only to find what disappeared, appeared, or moved.
4. Write each change with its printed page number as you go.

### 4. Build the year
1. Copy `fill_year_template.py` to `scripts/fill<YEAR>.py`; set SRC/DST/YEAR/PREV_YEAR.
2. Add every change to `O` (row → code) and every new agency to `NEW`, each with a page-cite comment.
3. Run it. Do not edit earlier years' cells — flag suspected hand-entry mistakes instead.

### 5. Check before delivering
1. `check_sheet.py <new.xlsx> <year> <previous.xlsx>` — must show no parent problems and 0 earlier-year
   cells changed.
2. Per-department check: every live child of a department must appear in that department's own list or
   officials page for this year (this catches carry-forward errors — the most common miss). Use
   `manual_text.py presence` for the whole sheet, then confirm each hit on the page; most are misspelled
   row names or renames.
3. Spot audit: pick ~20 random live rows and verify each against the page image; report the error rate
   and fix the cause, not just the row.

### 6. Compare with hand-entered data (when available)
1. `compare_sheets.py <mine> <hand workbook> "<tab>" <years...>`. Report the *live rows* agreement.
2. Classify every difference before calling it an error: **reading error** (mine or hand), **convention**
   (parent level, renames, blanks), **edition** (June vs. December), **not entered yet**.
3. Check every existence disagreement (one side Dead/not listed, the other live) on the page image. Fix
   your own errors; list the hand errors with page numbers and explanations — never silently copy the
   hand sheet, and never change the user's earlier years.
4. If the user asks for an error table, write an xlsx with: side, row, agency, my coding, hand coding,
   what the manual shows, printed page(s), explanation, wrong in the other edition too?

### 7. Deliver
- Save to `data/Spreadsheet 1935-<YEAR>.xlsx` in the repo and copy to `~/Downloads`; send/open the file.
- Report: edition used, totals by category, main changes, audit result, and a short list of judgment
  calls/open conventions with page numbers for the user to decide.
- Git: commit with a clear message. `git fetch` first; if the branch is behind or the last commit was
  already pushed, add a new commit — never amend or force-push published history. If git cannot
  authenticate, open GitHub Desktop (`open -a "GitHub Desktop" <repo>`) and tell the user to click
  **Push origin**.

## Coding at a glance (details in references/coding-rules.md)
- Cabinet: `Yes No None None No`; child of department: `No Yes <Dept> None No`; child of other:
  `No Yes None <Parent row name> No`; other: `No No None None Yes`; not listed: `No No None None No`;
  gone: `Dead` ×5.
- Units = chart boxes, "as follows" list entries, or own section headings — not officials.
- Executive Office is a parent; OEM members are OEM children; dashed line = no parent.
- International and quasi-official agencies: blank. Legislative/judicial: listed, blank, hidden.
  Dead rows: hidden, never deleted; unhide and relabel if they reappear.
