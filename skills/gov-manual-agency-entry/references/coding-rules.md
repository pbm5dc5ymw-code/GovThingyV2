# Coding rules (codebook)

Rules for turning one U.S. Government Manual edition into one year of the spreadsheet.
Sources: the workbook's **Instructions** and **FAQs** tabs (authoritative), and lessons from comparing
machine entry against hand entry for 1939-1943. Where the user has not yet decided a convention it is
marked **OPEN** — apply the default shown and list affected rows in the ambiguity list.

## Contents
1. Sheet layout
2. The five columns
3. What counts as a unit (a row)
4. Choosing the parent
5. Births, deaths, renames, reappearances
6. What to leave out and hide
7. Which edition is "the year"
8. Open conventions

## 1. Sheet layout
- Row 2 holds headers: `agency_name`, `agency_type`, then five columns per year:
  `<YEAR>_cabinet_agency`, `<YEAR>_child_agency`, `<YEAR>_cabinet_parent_agency`,
  `<YEAR>_other_parent_agency`, `<YEAR>_other_agency`. Find columns by header text, not by position.
- Years: 1935, 1936, 1937, 1939, 1940, ... There is no 1938 manual.
- `agency_type` is `Executive` for executive agencies; legislative/judicial rows say so and get no year values.
- New agencies are appended at the bottom. Never reorder or delete rows. Earlier years of a new row are
  "not listed" (`No / No / None / None / No`).
- Never change earlier years' cells (especially the user's hand-entered 1935-1939) — flag suspected
  mistakes instead.

## 2. The five columns
| Situation | cabinet | child | cabinet_parent | other_parent | other |
|---|---|---|---|---|---|
| Cabinet department | Yes | No | None | None | No |
| Child of a cabinet department | No | Yes | *Department name* | None | No |
| Child of a non-cabinet agency | No | Yes | None | *Parent's exact row name* | No |
| Standalone (other) agency | No | No | None | None | Yes |
| Not listed yet | No | No | None | None | No |
| Gone | Dead | Dead | Dead | Dead | Dead |

Cabinet departments 1935-1946: State, Treasury, War, Justice, Post Office, Navy, Interior, Agriculture,
Commerce, Labor. "Department of ..." is always cabinet; ask before treating anything else as cabinet.
Use the exact department strings already in the sheet (e.g. `Department of the Treasury`, `Department of War`).
A parent named in `other_parent` must exactly match a live row's `agency_name` that year.

## 3. What counts as a unit (a row)
Count it if it appears as **any** of:
- a box on that year's organization chart,
- an entry in the department's "principal bureaus / divisions / offices ... are as follows:" list,
- its own section heading with officials (and usually a "Creation and Authority" paragraph).

Do **not** count:
- officials or job titles (e.g. "Director of Personnel", "Adviser on International Economic Affairs"),
- committees or divisions only mentioned in passing,
- field/regional offices, branches inside a division (unless the user asks for that depth).

## 4. Choosing the parent
Decide from the **page image**, in this order:
1. The manual's reorganization tables / organization charts (solid line = authority; dashed = liaison only).
2. Creation text: "established within ...", "under the direction and supervision of ...".
3. The section it is printed in (including a department's "Related Organizations").
4. Running page headers — last resort.

Specific rules learned:
- A unit drawn **inside** another unit's box, or indented under it in a list, is that unit's child
  (U.S. Employment Service inside the Social Security Board; HOLC/FSLIC inside the FHLB Board).
- Boxes headed by an **official** (Assistant Secretary, Fiscal Assistant Secretary, Assistant to the
  Attorney General) are not parents — their units go to the department — **except** when a named
  organizational unit groups them (Fiscal Service → Bureau of Accounts, Bureau of the Public Debt, Treasurer).
- A transfer order moves the **top** agency; its subsidiaries keep their immediate parent
  (RFC subsidiaries stayed under the RFC when the RFC moved to Commerce in 1942).
- The Executive Office of the President **is** a parent (Budget Bureau, White House Office, OEM, Liaison Office, ...).
  The EOP itself is coded "other".
- Agencies the manual lists as "established within the Office for Emergency Management" are OEM children.
- An agency "related" to a department by a dashed line only (e.g. Commission of Fine Arts → Interior) stays standalone.
- Footnotes on charts can override the drawing (1942: starred agencies "treated as within the EOP").

## 5. Births, deaths, renames, reappearances
- **Dead** needs evidence: an Appendix A entry (abolished/transferred/consolidated) or absence from **both**
  that year's chart and unit list. Write "Dead" in all five columns for that year and every later year.
- **In liquidation**: if it still has its own section, keep it live and flag; if only mentioned as being
  liquidated elsewhere, treat as Dead and flag.
- **Reappearance** (FAQ): if a dead agency reappears, unhide the row and code its new status.
- **Merged into a new agency**: old rows Dead; new agency gets its own row.
- **Renamed** — OPEN (see 8). Default: if the sheet already has a row with the new name, old row Dead and
  use the new row; otherwise keep the existing row.
- **Duplicate rows** for the same agency: give them identical values unless the user has chosen one row.

## 6. What to leave out and hide
- **International agencies**: not considered (FAQ) — leave their year columns blank.
- **Quasi-official agencies** (the manual's own "Quasi-Official Agencies" section from 1941 on: National
  Academy of Sciences, Pan American Union, American National Red Cross): not considered — leave blank.
- **Legislative and judicial**: list them, leave year columns empty, hide the row.
- **Hide** (never delete) rows whose current-year status is Dead.

## 7. Which edition is "the year"
- FAQ: if a year has several manuals, use the **most recent** one dated in that year.
- Check govinfo for every edition first:
  `https://api.govinfo.gov/published/<year>-01-01/<year>-12-31?offset=0&pageSize=100&collection=GOVMAN&api_key=DEMO_KEY`
  PDFs: `https://www.govinfo.gov/content/pkg/<packageId>/pdf/<packageId>.pdf`
- Known editions: 1939-10-01, 1940-09-01, 1941-09-01, 1942-03-01 and 1942-09-01 (use Sept), 1943-06-01 and
  1943-12-01 (Dec used; OPEN), 1944-06-01, 1945-01-01 and 1945-02-01.
- From the Winter 1943-44 edition on, organization charts were discontinued — use unit lists and officials pages.

## 8. Open conventions (ask the user; until then use the default and flag)
| Question | Default used so far | Hand-entered sheet tends to |
|---|---|---|
| Parent level: immediate unit or larger agency? | Immediate unit as printed (SSB, OEM, FHLB Administration) | Skip to the larger agency (FSA, EOP, NHA) |
| Renamed unit: new row or same row? | Same row unless a new-name row already exists | Old row Dead, new row live |
| Pre-set blank rows for units that exist | Fill them | Fill them |
| Which 1943 edition | December 1943 | Mixed June/December |
| Joint Chiefs of Staff as a row (OSS parent) | Added as "other" | — |
