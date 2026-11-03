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
- its own section heading with officials (and usually a "Creation and Authority" paragraph),
- a centered small-caps subheading inside another agency's section (e.g. REGIONAL AGRICULTURAL CREDIT
  CORPORATIONS inside Farm Credit Administration) — it is a child of that agency. This applies to
  agencies that have **no** "as follows" unit list; inside a cabinet department the department's own unit
  list decides (subheadings in its Activities section, e.g. Navy's JUDGE ADVOCATE GENERAL under Executive
  Offices of the Secretary, are sub-units of a listed unit and are not added),
- (a "Chief/Director, <unit>" line in an officials list only shows WHERE a unit sits; it is not
  evidence by itself — see below).

**Format comes first: a set-off name followed by a description** (user rule, 1949-50 review). Each live row
needs, in that year's chapter text, its name set off typographically (bold/centered heading or SMALL-CAPS
run-in "OFFICE OF X.—This office ...") and then a description of the unit. Appendix, charts, unit lists and
branch lists support a decision but do not qualify a unit alone. Example: "National Archives Trust Fund Board"
(bold, then a description) is in; "Division of Transportation Accounts" (named inside a sentence) is out.
What "set off" looks like differs by edition, so check a few pages of each edition first:
- 1944-48 (one column): bold centered headings, smaller centered capitals inside a section ("GENERAL BOARD OF
  THE NAVY"), or a paragraph opening in capitals and small capitals ("Division of Traffic.—This Division ...").
  A State geographic division's description may be only its territory ("Division of African Affairs.—Ethiopia,
  Liberia, ..."); that counts. Several units named in one run-in count.
- 1949: same headings, with run-ins in full small capitals.
- 1950 (two columns): bold headings and "OFFICE OF X.—" run-ins; Divisions under Offices are plain text.
These are NOT enough on their own:
- a job title in an officials list ("Director, Division of College Housing");
- a box on an organization chart (charts decide parents, not existence);
- a plain, not set-off paragraph describing an official ("The Director of Management Research is responsible ...");
- a program/activity paragraph ("COLLEGE HOUSING.—The Housing Act of 1950 authorizes the Administrator");
- a name that is only mentioned in plain body text, not set off typographically ("The International
  Business Practices Policy Staff discharges ..." in ordinary type) — check the page image;
- a name that appears only in a list or sentence naming several units ("comprises a Division of X and a
  Division of Y"), or only in a directory table with a heading (U. S. Attorneys and Marshals, 1950);

The user's three requirements (2026-11-03), all needed:
1. The set-off name is an **agency, not a position**. "OFFICE OF THE CHIEF OF ENGINEERS.—", "CORPS OF
   ENGINEERS.—", "Office of the Solicitor General" pass. "CHIEF OF ENGINEERS.—", "THE ATTORNEY GENERAL.—",
   "LEGAL ADVISER.—", "CHIEF, NATIONAL GUARD BUREAU.—", a bold "Director of Liquidation" or "Secretary of the
   Navy" fail, even with a description. One exception kept: a bold title whose text describes the Office as an
   organization ("The Chief of Naval Operations" / "The Office of the Chief of Naval Operations is organized ...").
2. **Formatting** sets the agency apart: usually bold; a capitals-and-small-capitals run-in also counts.
3. A **description** of the agency follows.
A name only in a unit list or officials list fails. Army service chiefs are judged this way, row by row and year
by year (1950 prints every chief as an official, so they are all off in 1950). Do not add units below the depth
1935-48 use, such as Navy boards or Veterans Administration offices, even if they pass.
Use scripts/format_crops.py to make contact sheets of every new or changed row and look at each one.

The same test applies to **every row every year** — rows carried from last year get no benefit of the
doubt. Record the evidence type and printed page for each live row; a row with no evidence is not live.

Do **not** count:
- officials or job titles (e.g. "Director of Personnel", "Adviser on International Economic Affairs"),
- committees, corporations or divisions only mentioned in a sentence of another unit's paragraph
  (1944: Federal Prison Industries, Defense Homes Corporation, Arboretum Advisory Council) — even if they
  still legally exist,
- run-in activity paragraphs ("FEDERAL CREDIT UNIONS.—Federal credit unions are ...") that describe a
  program, not an office,
- field/regional offices, branches inside a division (unless the user asks for that depth).

## 4. Choosing the parent
Decide from the **page image**, in this order (heading level first: a unit with its own top-level
heading is standalone unless the page says "established within" / "under the administration of"):
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
- **Duplicate rows** for the same agency: keep live the row whose name matches the manual's wording
  ("Bureau of the First Assistant Postmaster General", not "Office of ...") and mark the other Dead —
  this is what the hand-entered sheet does. If neither matches, give them identical values.

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
  1943-12-01 (Dec used; OPEN), 1944-06-01 (only 1944 edition; text runs to Aug 1944), 1945-01-01 and 1945-02-01.
- From the Winter 1943-44 edition on, organization charts were discontinued — use unit lists and officials pages.

## 8. Open conventions (ask the user; until then use the default and flag)
| Question | Default used so far | Hand-entered sheet tends to |
|---|---|---|
| Parent level: immediate unit or larger agency? | Immediate unit as printed (SSB, OEM, FHLB Administration) | Skip to the larger agency (FSA, EOP, NHA) |
| Renamed unit: new row or same row? | Same row unless a new-name row already exists | Old row Dead, new row live |
| Pre-set blank rows for units that exist | Fill them | Fill them |
| Which 1943 edition | December 1943 | Mixed June/December |
| Joint Chiefs of Staff as a row (OSS parent) | Added as "other" | — |
