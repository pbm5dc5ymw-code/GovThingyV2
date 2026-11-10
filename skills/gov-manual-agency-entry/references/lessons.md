# Lessons from 1939-1943 (why entries went wrong)

Every miss found when machine entry was compared with hand entry traced to process, not to reading
ability. Keep these in mind; each has a fix built into the SKILL.md workflow.

| What went wrong | Example | Root cause | Fix |
|---|---|---|---|
| Rules taken from one example | "Dead stays Dead", "EOP is never a parent" | Treated the sample sheet as the full spec; didn't read the Instructions/FAQs tabs | Read every tab first; restate rules; ask before generalizing from one example |
| Wrong edition | March 1942 used instead of Sept 1942 (~50 rows) | Used whatever PDF was downloaded | Check govinfo for every edition of the year; use the latest |
| Summaries skipped | 1939: Federal Prison Industries, Foreign Service Buildings Commission placed wrong | Never read the reorganization tables / Appendix A | Read reorg tables, Appendix A, charts before departments |
| Text instead of layout | Post Office divisions, Cost Ascertainment under wrong parent | Extracted text merges two-column lists and drops indentation | Decide structure from rendered page images |
| Carry-forward | 1943: Navy JAG, State Personnel Supervision, FHLB Board kept alive | Defaulted to last year's coding | Rebuild each department from its own list every year; run the per-department check |
| Official treated as unit | "Director of Personnel", "Adviser on International Economic Affairs" | Name looked like an office | Only list entries, chart boxes, section headings count |
| Transfer target used as parent | FNMA, Disaster Loan put under Commerce in 1942 | Read appendix transfer literally | Transfer moves the top agency; subsidiaries keep their parent |
| Copying a reference sheet against the page | Social Security bureaus and FHLB units switched to match feedback, then wrong | Trusted the hand sheet over the layout | Layout wins; record disagreements instead of copying |
| Exact-name search | Misspelled row names ("Chemsity", "Adminstration") looked missing | Literal string matching | Fuzzy/partial matching; confirm absence on the page image |
| Chart-era parent kept in chartless years | 1943: OPA and Office of Censorship left under the EOP after charts stopped | A 1942 chart footnote ("treated as within the EOP") was carried forward; the Dec 1943/1944 EOP list on p. 53 omits them | Re-read the EOP (p. 53) and OEM (p. 61) membership lists every year; an agency off both lists is standalone |
| Units missing from the sheet entirely | Office of the Solicitor General, OASI bureau, National Capital Housing Authority, WFA offices | Checked only existing rows against the manual, never the manual's list against the rows | Walk each unit list item by item and confirm a row exists (search dead rows too — reuse them if the unit reappears) |
| Inclusion test applied unevenly | 1944: killed State boards mentioned in passing but kept Federal Prison Industries, Defense Homes Corp. (also one sentence) | Rows carried from last year were never re-tested | Evidence table for every row each year (evidence.py) |
| Parent by history instead of heading level | 1944: National Archives Council under National Archives though it has its own top-level heading | Followed earlier coding / chairmanship | Heading level first; only "within"/"under the administration of" makes it a child |
| Blank rows never checked | Cooperative Research and Service Division (FCA subheading) left blank | Checked only live rows | Run the evidence check on blank and dead rows too |
| Duplicate rows both live | Office vs Bureau of the First Assistant Postmaster General | No rule | Keep the row matching the manual's wording; other Dead |
| Job titles and chart boxes taken as units | 1950: HHFA Divisions of College Housing, Housing Research, Slum Clearance; Interior admin divisions; Veterans' Employment Service; DoD Staff Council | Treated an officials line or chart box as printed evidence | A unit needs a heading, list entry, or a description of the unit itself (evidence.py now ignores chart pages and treats OFFICIAL as weak) |
| Names in plain body text taken as units | 1950: State economic "Staffs", Division of International Administration, Army "CHIEF OF ..." run-ins | Text search found the name; layout never checked | Contact-sheet review of every new/changed row (format_crops.py); unit must be typographically set off |
| Rewrote a pushed commit | Amended a commit the user had already pushed | Didn't check the remote | `git fetch`; if behind or diverged, add a new commit instead of amending |

## Format-priority review of 1949-50
- The same units can change format between editions: State and Post Office divisions are run-ins in 1949
  but plain sentences or lists in 1950. Judge each edition from its own pages; do not carry a verdict over.
- The OCR text layer has no bold and often truncates headings, so "no match" is not "no heading". Before
  removing a row, check the current name: rows keep old names after renames (row "Office of Field Operations"
  is printed "Office of Field Service"). Three removals were caught this way on re-check.
- Keep additions at the depth used in earlier years; sub-bureau run-ins (Internal Revenue units, Public Debt
  divisions, VA offices) qualify by format but would make 1949-50 inconsistent with 1935-48. Ask first.

## Format-priority review of 1944-48
- Hand-entered cells can carry trailing spaces ("Yes "). Strip before testing a code, or live agencies look
  unlisted (NLRB and the Smithsonian looked unlisted until this was caught).
- Agency names repeat across departments ("Bureau of Accounts", "Office of the Secretary"). Match a heading only
  inside the unit's own department pages, using the running page headers for the range, and treat a parent as
  gone only if no row of that name is live that year.
- Review both directions: live rows with no set-off text, and dead or unlisted rows that are set off in print.

## Where hand-entered data tends to slip (useful when comparing)
- Agencies that ended in earlier years left alive (renamed or abolished 1-2 years before).
- Existing agencies marked Dead (often units inside another agency's section).
- Whole departments' sub-units left blank or coded "not listed" (War Department in 1943).
- Mixing two editions of the same year.
Treat hand differences as hypotheses to check on the page, not as errors to copy or ignore.

## Agency, not position (2026-11-03)
Earlier passes accepted a set-off title with a description ("THE INSPECTOR GENERAL.—The Inspector General ...").
The user's rule is that the set-off name itself must name an agency (Office, Bureau, Corps, Department, Division,
Board, Staff ...). Scan every edition for run-ins and bold headings that are titles (Chief of, Director of, The X
General, Adviser, Secretary) and check which rows rest on them. Also re-check "Office of the Secretary" rows:
in 1948-49 several were only in unit lists or officials lists.

## 1951 (2026-11-03)
The 1951-52 edition is two columns like 1950, but Interior and Commerce now head the Secretary's staff with bold job
titles, so staff divisions that were rows in 1950 drop out. Compare every live row to the new edition's headings,
then sweep each chapter's headings for units with no row: this found units printed in 1950 that never got rows
(Air Force major commands, Post Office staff offices, Postal Savings System, Weapons Systems Evaluation Group).
Scratchpad tools used: inv2.py (two-column heading/run-in inventory), cmp51.py, dept.py (per-chapter comparison).

## 1952 (2026-11-03)
The 1952-53 scan skips printed pp. 386-387 and its offset changes (+12 for pp. 57-385, +10 after): map pages before
rendering. Commerce now describes its staff offices in plain paragraphs under bold Assistant Secretary sections, so
they end. Look at every chapter's page images for bold "Office of ..." headings, not only the rows the matcher flags:
the 1951 Navy Offices of Naval Material and Naval Research were missed that way and filled in 1952.

## 1953 (2026-11-03)
The 1953-54 scan has a constant offset (+9). New executive department (HEW): give it a new row and end the Federal
Security Agency row, because Reorganization Plan 1 of 1953 abolished the FSA rather than renaming it. Every year,
also sweep the dead and unlisted rows against standalone heading lines in the new text (a line that is exactly
the row's name, followed by CREATION/PURPOSE or "The ... is"). That sweep and the HHFA/GSA page images showed three
rows wrongly coded Dead since 1950-51: Federal Home Loan Bank System, Federal National Mortgage Association, and
National Archives Trust Fund Board. Their bold headings had been cut short in OCR ("Federal National M"), and the
earlier passes read the cut-short text as "sentences only". Confirm on the image before calling a unit unprinted.

## 1954 (2026-11-10)
The 1954-55 scan has a constant offset (+8). The scratchpad (venv and helper scripts) can be wiped between
sessions; rebuild with pymupdf, openpyxl and pillow. The pymupdf text of this edition splits lines differently,
so exact-line heading matches miss many real headings: search the flattened page text for "Name The ..." or
"Name CREATION" and confirm on the page image. For a renamed bureau whose children already name it as parent
(e.g. the State consular bureau), keep the row name so the children's parent cells stay valid.

## 1955 (2026-11-10)
The 1955-56 edition (revised June 1, 1955) has a constant offset (+8). A unit can drop to a plain list one year
and come back as a run-in the next (Commerce staff offices 1952-54 vs 1955; State's Foreign Service Institute and
Inspection Corps): reuse the dead row rather than adding a duplicate. Search dead rows by name before adding.
