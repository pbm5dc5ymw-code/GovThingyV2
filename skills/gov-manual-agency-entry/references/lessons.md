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
| Rewrote a pushed commit | Amended a commit the user had already pushed | Didn't check the remote | `git fetch`; if behind or diverged, add a new commit instead of amending |

## Where hand-entered data tends to slip (useful when comparing)
- Agencies that ended in earlier years left alive (renamed or abolished 1-2 years before).
- Existing agencies marked Dead (often units inside another agency's section).
- Whole departments' sub-units left blank or coded "not listed" (War Department in 1943).
- Mixing two editions of the same year.
Treat hand differences as hypotheses to check on the page, not as errors to copy or ignore.
