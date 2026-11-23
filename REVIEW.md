# Technical review — Financial Performance (P&L)

**Project:** `Financial_Performance_PnL.pbip` (Report + SemanticModel, PBIR + TMDL)
**Date:** 2026-11-13

This is a review of the project against its own first working build, followed by the changes made in
response and the verification of those changes. It is kept in the repository because the numbers
section is an independent check of every figure the report puts on screen, and because several of
the decisions below are deliberate and would otherwise look like oversights.

## How this was verified

- **Model metadata** was read from the TMDL definition and loaded into an offline tabular model.
- **Numbers** were recomputed independently in Python straight from the source workbook the model
  loads (`data\PnL_Dataset.xlsx`), replicating each measure's DAX filter logic row by row. The model
  is a plain Import of that workbook with no calculated columns or tables, so this is equivalent to
  querying the model — and a stronger check, because it does not reuse the model's own logic.
- **BPA** was run with Tabular Editor 2.29.0 against a `.bim` export of the model, read-only.
- After the changes, every figure was re-checked against the live model in Power BI Desktop.

---

## Summary

The model was correctly shaped from the start: a single fact, five conformed dimensions, five
many-to-one single-direction active relationships, no bidirectional or many-to-many anywhere, no
RLS, auto date/time off. Every division used `DIVIDE` with an explicit alternate result. The
narrative numbers on each page were correct to the digit at the default slicer state.

Three things needed work.

**First, the headline KPI on the Budget vs Actuals page was a data artefact, not an insight.** The
card showed Net Profit actual 46,956 against budget 27,897, a **+68.3%** beat, driven entirely by
the way budget OPEX is generated: all five OPEX accounts receive an identical budget total, while
actual OPEX is generated with realistic per-account weights. Revenue and COGS are generated
sensibly (actual/budget 0.978 and 1.010); OPEX is not. The largest number on the page was noise,
and the page did not say so. See §3.6.

**Second, "Net Profit" was EBIT under another name.** The fact table carries no interest rows, so
the interest measures were always blank and the net profit measures were arithmetically identical
to EBIT.

**Third, formatting was incomplete and two display bugs were visible.** 24 of 29 measures had no
format string, no measure had a description, and the COGS ratio chart rendered `0.54` / `0.48`
where the text beneath it said "54% vs 48%".

Everything else was polish: 10 of 29 measures unused, 19 columns dead weight, foreign keys visible
in the field list, and `Dim_Date` unable to be marked as a date table because its `Date` column was
typed as a whole number.

---

## 1. The baseline

### 1.1 Data sources and loading

All six source tables loaded from **one hard-coded absolute path** holding the Windows user name,
with **no Power Query parameters** — so changing the source meant editing six queries, and the user
name would have entered the repository on the first push.

Each query was the same four steps: `Excel.Workbook` → sheet navigation → `Table.PromoteHeaders` →
`Table.TransformColumnTypes`, with explicit typing on every column. Two generated tables existed:
`Last Refreshed Date` (`DateTime.LocalNow()`, 1 row) and `Calendar` (a 1,200-day `List.Dates`
calendar with 20 derived columns).

### 1.2 Tables and row counts

| Table | Rows | Visible | Notes |
|---|---|---|---|
| `Fact_Financials` | 50,688 | yes | 25,344 Actual + 25,344 Budget |
| `Dim_Date` | 366 | yes | 2024 only, daily grain |
| `Dim_Account` | 16 | yes | **never used by any visual** |
| `Dim_Brand` | 11 | yes | |
| `Dim_Region` | 4 | yes | |
| `Dim_Channel` | 4 | yes | |
| `Calendar` | ~1,200 | **hidden** | unused, **no relationship** |
| `Last Refreshed Date` | 1 | **hidden** | |
| `_Measures Table` | 0 | yes | measure holder |

Fact rows by category: Revenue 12,672 / COGS 16,896 / Operating Expenses 21,120. **No
`Interest Expense` rows at all** — see F-3.

Grain mismatch: the fact had **12 distinct `DateKey` values** (one per month of 2024) against 366
daily rows in `Dim_Date`. Harmless for the visuals in use, but any day or week visual would be empty.

### 1.3 Relationships

Five relationships, **all many-to-one, single-direction, active**.

| From | To | Cardinality | Direction |
|---|---|---|---|
| `Fact_Financials[DateKey]` | `Dim_Date[DateKey]` | Many→One | Single |
| `Fact_Financials[BrandKey]` | `Dim_Brand[BrandKey]` | Many→One | Single |
| `Fact_Financials[RegionKey]` | `Dim_Region[RegionKey]` | Many→One | Single |
| `Fact_Financials[ChannelKey]` | `Dim_Channel[ChannelKey]` | Many→One | Single |
| `Fact_Financials[AccountName]` | `Dim_Account[AccountName]` | Many→One | Single |

All five were named `AutoDetected_<guid>` — created by auto-detect rather than modelled
deliberately. The `Dim_Account` join used the **text column `AccountName`** while an integer
`AccountKey` sat unused.

### 1.4 The date table

- **Not marked as a date table**, and **could not be marked as it stood**: `Dim_Date[Date]` was
  typed `Int64`, holding Excel serials (`45292`, `45293`, …). Marking requires a unique, contiguous,
  date-typed column, so **no DAX time intelligence worked** — `SAMEPERIODLASTYEAR`, `DATEADD` and
  `TOTALYTD` would all have failed or misled.
- Auto date/time was off (`__PBI_TimeIntelligenceEnabled = 0`). Correct.
- `MonthName` correctly carried `sortByColumn: Month`.

### 1.5 Measures (29)

The cascade was complete — Revenue → COGS → Gross Profit → OPEX → EBIT → Interest → Net Profit —
organised into display folders (CORE, GROSS PROFIT, NET PROFIT, VARIANCE, COGS RATIO, KPI CARD
LABELS). Every division used `DIVIDE` with an explicit `0` alternate; no raw `/` anywhere.

**No measure had a description. 24 of 29 had no format string.**

### 1.6 Calculated columns, tables and roles

No calculated columns — all derivation in Power Query. No calculated tables, no RLS, no calculation
groups, perspectives, hierarchies or translations.

---

## 2. Findings

### 2.1 DAX

**F-1 — Ratio measures had no percentage format string, and two were on screen.** `Actual COGS
Ratio`, `Budget COGS Ratio` and `COGS Ratio Variance` carried only
`PBI_FormatHint = {"isGeneralNumber":true}`. The COGS ratio bar chart had data labels enabled, so it
rendered `0.54` / `0.48` directly above text reading "54% vs 48%".

**F-2 — `Gross Margin vs Budget Label` rendered the wrong number.**

```dax
VAR v = [Actual Gross Margin %] - [Budget Gross Margin %]
RETURN IF(v >= 0, "▲ " & FORMAT(v, "+0.00pp"), "▼ " & FORMAT(ABS(v), "0.00pp"))
```

In a DAX custom numeric format string `%` is the specifier that scales by 100; `p` is not a
specifier and is emitted literally. So `FORMAT(0.01341, "0.00pp")` yields `"0.01pp"`, not
`"1.34pp"`. The card read **`▼ 0.01pp`** where the gap was −1.341pp. Fix:
`FORMAT(ABS(v) * 100, "0.00") & "pp"`.

**F-3 — `Actual Net Profit` was EBIT, mislabelled.** With no interest rows in the fact, the interest
measures were always blank and `Actual Net Profit` ≡ `Actual EBIT` = 46,956.04, `Budget Net Profit`
≡ `Budget EBIT` = 27,897.43. `Dim_Account` defined `Interest on Debt` plus `Gross Profit`, `EBIT`
and `Net Profit` as calculated rows, none of which appeared in the fact.

**F-4 — The variance sign convention was arithmetically consistent but semantically inconsistent.**
Every variance was `Actual − Budget`, so positive meant favourable for revenue and gross profit but
unfavourable for COGS, OPEX and the COGS ratio. The summary table placed `Revenue vs Budget Label`
(▲ = good) beside `COGS Variance` as a raw signed number, so within one row a positive number meant
good in one column and bad in the next, with nothing to tell the reader.

**F-5 — Budget vs Actual logic was correct at every grain.** 0 Actual-version rows carried a
non-zero `BudgetAmount` and 0 Budget-version rows a non-zero `ActualAmount`. Additivity held across
categories, quarters and months, and percentage measures were correctly recomputed per grain rather
than summed.

**F-6 — Four implicit measures.** The "last refreshed" card on every page used
`Min(Last Refreshed Date[Last Refreshed])` — a column aggregation dropped into the visual rather
than a measure. `discourageImplicitMeasures` was not set.

**F-7 — No descriptions, and most measures unformatted.** Currency measures carried no format
string, so revenue displayed as a bare `91,077` with no symbol and no units anywhere in the report.
`Dim_Region` carried a `Currency` column that was never used.

**F-8 — A placeholder `Example Measure = 0` was still in the model and visible,** with an `en-AU`
currency format hint.

**F-9 — 10 of 29 measures were unused by any visual,** including conditional formatting and sort
references. Several were legitimate building blocks (`Actual COGS` feeds Gross Profit) and were kept.

### 2.2 Model

**M-1 — No bidirectional or many-to-many relationships.** All five many-to-one, single-direction,
active.

**M-2 — No key type mismatches.** Integer joined to integer, text to text.

**M-3 — Foreign and primary keys were visible in the field list.** 10 key columns unhidden, and
`FactKey`, `BudgetAmount` and `ActualAmount` visible and summarizable — so a user could drag
`ActualAmount` into a visual and get a number that ignored the `VersionType` and `AccountCategory`
logic entirely. No dimension key was marked as a primary key.

**M-4 — Large amount of unused columns.** Only **6 columns** were used in any field well across the
whole report. Dead weight included the entire `Calendar` table (20 columns, ~1,200 rows, hidden,
unrelated, regenerated on every refresh because it was anchored to `DateTime.LocalNow()`), the
entire `Dim_Account` table, and 11 columns of `Dim_Date`.

**M-5 — `Fact_Financials[AccountCategory]` duplicated `Dim_Account[AccountCategory]`.** Every
measure filtered the fact's copy rather than the dimension's. The two agreed on every row, so
nothing was broken — but the denormalised column was why `Dim_Account` was unused, and why the
by-account chart pulled its axis from the fact table instead of the dimension.

**M-6 — `Dim_Date` was not, and could not be, marked as a date table.** See §1.4.

**M-7 — `BudgetAmount` and `ActualAmount` were `double`.** Floating point on currency columns;
`decimal` is the safer choice.

### 2.3 Best Practice Analyzer — baseline

Tabular Editor 2.29.0, standard rule set, read-only: **207 violations across 18 rules.**

| Count | Rule |
|---|---|
| 81 | [Maintenance] Visible objects with no description |
| 24 | [Formatting] Whole numbers should be formatted with thousands separators and no decimals |
| 24 | [Formatting] Provide format string for measures |
| 19 | [Performance] Set IsAvailableInMdx to false on non-attribute columns |
| 19 | [Maintenance] Remove unnecessary columns |
| 10 | [Formatting] Hide foreign keys |
| 5 | [Formatting] Mark primary keys |
| 4 | [Formatting] Percentages should be formatted with thousands separators and 1 decimal |
| 3 | [Performance] Date/calendar tables should be marked as a date table |
| 3 | [Maintenance] Ensure tables have relationships |
| 3 | [Formatting] Month (as a string) must be sorted |
| 3 | [Formatting] Do not summarize numeric columns |
| 2 | [Performance] Do not use floating point data types |
| 2 | [Formatting] Provide format string for "Month" columns |
| 2 | [Formatting] Hide fact table columns |
| 1 | [Performance] Remove redundant columns in related tables |
| 1 | [Performance] Model should have a date table |
| 1 | [Performance] Minimize Power Query transformations |

Of the 207, about **120 were the `Calendar` table alone**.

---

## 3. Numbers

Recomputed from the source workbook (50,688 fact rows), replicating each measure's filter logic.
Default slicer state = fully unfiltered.

### 3.1 Totals

| Measure | Value |
|---|---|
| Actual Revenue | **91,077.47** |
| Budget Revenue | **93,145.92** |
| Revenue Variance | **−2,068.45** |
| Revenue Variance % | **−2.221%** |
| Actual COGS | 38,427.51 |
| Budget COGS | 38,050.89 |
| Actual COGS Ratio | **42.19%** |
| Budget COGS Ratio | **40.85%** |
| Actual Gross Profit | 52,649.96 |
| Budget Gross Profit | 55,095.03 |
| Gross Profit Variance | −2,445.07 |
| Actual Gross Margin % | **57.808%** |
| Budget Gross Margin % | **59.149%** |
| Gross Margin gap | **−1.341pp** |
| Actual OPEX | **5,693.92** |
| Budget OPEX | **27,197.60** |
| Actual EBIT | 46,956.04 |
| Budget EBIT | 27,897.43 |
| EBIT Variance % | **+68.317%** |

### 3.2 By brand category

| Category | Actual Rev | Budget Rev | Rev Var % | COGS ratio A / B | GM % A / B |
|---|---|---|---|---|---|
| Consumer Beauty | 24,309.04 | 27,603.42 | **−11.935%** | **54.20% / 48.00%** | 45.805% / 52.000% |
| Prestige Fragrance | 33,627.30 | 32,832.03 | **+2.422%** | 35.00% / 35.00% | 65.000% / 65.002% |
| Professional Beauty | 11,362.77 | 11,308.92 | +0.476% | 42.00% / 42.00% | 58.001% / 57.998% |
| Skincare | 21,778.36 | 21,401.55 | +1.761% | 40.00% / 40.00% | 60.000% / 59.999% |

### 3.3 By region

| Region | Actual Rev | Budget Rev | Rev Var % | GM % |
|---|---|---|---|---|
| Asia Pacific | 23,738 | 22,740 | **+4.390%** | 58.06% |
| Europe | 27,265 | 27,968 | −2.515% | 58.30% |
| Latin America | 10,819 | 12,160 | −11.031% | 56.45% |
| North America | 29,256 | 30,278 | −3.375% | 57.65% |

### 3.4 By channel — Revenue Variance

| Channel | Rev Var | Rev Var % |
|---|---|---|
| Retail | **−864** | −2.318% |
| E-Commerce | −514 | −2.207% |
| Wholesale | −503 | −2.159% |
| Direct | −188 | −2.017% |

### 3.5 Additivity and grain consistency

All pass:

- `Budget Revenue + Revenue Variance = Actual Revenue` — exact
- Sum of the 4 categories = total, for Actual Revenue (91,077.47), Budget Revenue (93,145.92) and
  Actual EBIT (46,956.04) — exact
- Sum of the 4 quarters = total Actual Revenue — exact: Q1 18,585 / Q2 21,480 / Q3 24,229 / Q4 26,784
- Sum of the 12 months = total Actual Revenue — exact
- Actual and Budget columns cleanly separated: 0 cross-contaminated rows in either direction
- `Fact_Financials[AccountCategory]` agreed with `Dim_Account[AccountCategory]` on every row; no
  orphan or duplicate `AccountName`; no orphan `DateKey`; no duplicate `DateKey` in `Dim_Date`

Quarterly `Revenue Variance %`: −2.170% / −2.294% / −2.077% / −2.327% against −2.221% total —
correctly recomputed per grain rather than averaged.

### 3.6 The OPEX allocation (the one real number issue)

Budget OPEX is distributed **perfectly flat** across the five OPEX accounts; actuals are not:

| Account | Actual | Budget | Actual/Budget |
|---|---|---|---|
| Marketing & Advertising | 2,141.94 | **5,439.52** | 0.394 |
| Salaries & Benefits | 1,648.09 | **5,439.52** | 0.303 |
| R&D | 815.85 | **5,439.52** | 0.150 |
| Office & Facilities | 652.75 | **5,439.52** | 0.120 |
| Depreciation | 435.29 | **5,439.52** | 0.080 |
| **Total OPEX** | **5,693.92** | **27,197.60** | **0.209** |

The other categories are generated sensibly:

| Account | Actual | Budget | Actual/Budget |
|---|---|---|---|
| Product Sales | 68,308.07 | 69,859.14 | 0.978 |
| Service Revenue | 13,661.76 | 13,971.51 | 0.978 |
| Licensing Income | 9,107.64 | 9,315.27 | 0.978 |
| Raw Materials | 13,449.55 | 13,317.75 | 1.010 |
| Manufacturing Labor | 11,528.25 | 11,415.49 | 1.010 |
| Packaging | 7,685.57 | 7,609.97 | 1.010 |
| Logistics | 5,764.14 | 5,707.68 | 1.010 |

All 10,560 Actual OPEX rows and all 10,560 Budget OPEX rows are non-zero and at the same grain
(11 brands × 4 regions × 4 channels × 5 accounts × 12 months), so this is not a missing-rows
problem — the budget generator simply gave every OPEX account an identical total. Actual OPEX is
6.25% of revenue; budget OPEX is 29.2%.

**Consequence:** the largest figure in the report — the EBIT beat against budget — is a
data-generation artefact. This is in the **source data**, not the DAX. The decision taken was to
surface it rather than hide it or quietly fix the data: see §4.2.

---

## 4. Changes applied

The model layer was driven through the Tabular Object Model and re-serialised to TMDL, so the
metadata is tool-generated rather than hand-written. Power BI Desktop was closed for the structural
edits — deleting a table, retyping `Dim_Date[Date]`, adding a Power Query parameter and every report
JSON edit — because those against a live model need a refresh mid-flight and can half-apply. Desktop
was reopened at the end to refresh and verify.

One note for anyone editing these files with a script: writing them with PowerShell
`Set-Content -Encoding utf8` produces **UTF-8 with BOM** under Windows PowerShell 5.1, and Power BI
refuses to open the project ("Only text with UTF8 encoding without BOM is supported"). Write them
byte-wise without a BOM.

### 4.1 Model

| Area | Change |
|---|---|
| Dead weight | `Calendar` deleted — hidden, unrelated, never used, regenerated on every refresh, and the source of roughly 120 of the 207 baseline BPA findings. `Example Measure` deleted. |
| Net Profit → EBIT | `Actual Net Profit`, `Budget Net Profit`, `Actual Interest Expense` and `Budget Interest Expense` deleted; `Net Profit Variance %` renamed `EBIT Variance %` and rebased on EBIT. The dataset has no interest rows, so EBIT is now openly the bottom line. |
| Date table | `Dim_Date[Date]` retyped from `Int64.Type` to `type date`, and `Dim_Date` **marked as the model date table** (`dataCategory: Time`, `isKey`, `isUnique`). Time intelligence now works; it could not before. |
| Star schema | Base measures now filter `Dim_Account[AccountCategory]` instead of the fact's duplicate, and `Fact_Financials[AccountCategory]` was deleted. `Dim_Account` went from unused to driving the P&L page. |
| Source | `SourceFolder` text parameter added; all six queries read `SourceFolder & "\PnL_Dataset.xlsx"`. The workbook moved into the project at `data\PnL_Dataset.xlsx` (verified byte-identical by SHA-256). Six hard-coded absolute paths and the Windows user name are gone from the committed TMDL. |
| Variance convention | One convention everywhere: **positive = favourable**. `COGS Variance`, `OPEX Variance` and `COGS Ratio Variance` replaced by `COGS Variance (Fav)`, `OPEX Variance (Fav)` and `COGS Ratio Variance (Fav)`, each `Budget − Actual`. `EBIT Variance` added. |
| P&L measures | `P&L Actual`, `P&L Budget`, `P&L Variance (Fav)` and `P&L Variance %` added. Leaf accounts sum the fact; the Gross Profit and EBIT rows clear the account filter and return the subtotal; the favourable sign is decided from `Dim_Account[IsCost]`, so a new cost account picks it up automatically. |
| Narrative measures | `Narrative Revenue`, `Narrative Margin` and `Narrative Budget vs Actuals` added in a `NARRATIVE` folder. |
| Formatting | All 33 measures carry a format string except the five returning text: currency `\€#,0;-\€#,0;\€0`, percentages `#,0.0%;-#,0.0%;#,0.0%`. Leftover `PBI_FormatHint` annotations stripped. |
| Descriptions | **All 33 measures, all 7 visible tables and all 17 visible columns** now have descriptions. There were none at all. |
| Columns | Deleted: `Fact_Financials[FactKey, AccountCategory]`, `Dim_Account[AccountKey, PLOrder, IsRevenue]`, `Dim_Date[MonthShort, Week, Day, Weekday, IsWeekend, YearMonth, FiscalQuarter]`. Added `Dim_Account[PLSort]` (= `PLOrder × 100 + AccountKey`, hidden) as the statement sort order, since `PLOrder` is a section number and not unique per account. |
| Keys and amounts | All 5 foreign keys and both amount columns hidden, `summarizeBy: none`, `isAvailableInMdx: false` on hidden non-sort columns. Four dimension keys marked as primary keys. `BudgetAmount` and `ActualAmount` retyped `double` → `decimal` (`Currency.Type`), which also removed the floating-point noise: totals are now exactly `91,077.47` rather than `91,077.47000000015`. |
| Sorting | `Dim_Date[Quarter]` sorts by `QuarterNum`; `Dim_Account[AccountName]` sorts by `PLSort`. |
| Housekeeping | `Last Refreshed Date` renamed `Refresh Stamp` (the old name tripped the date-table rule on a non-calendar table); the five `AutoDetected_<guid>` relationships renamed `Fact_Financials_Dim_<Dimension>`; stale Q&A linguistic metadata referencing deleted tables reset. |

**Model size: 3,177,376 → 1,027,406 bytes (−68%).** 9 tables → 8, 29 measures → 33.

### 4.2 Report

| Area | Change |
|---|---|
| Residue | An unused placeholder page deleted outright rather than left hidden. All 6 orphan bookmarks deleted — they targeted three page IDs that no longer existed. An unused `textFilter` custom visual removed from `publicCustomVisuals`, dropping an external dependency from a published report. |
| Theme | A neutral, colour-blind-safe palette (`NeutralCBSafe.json`), five of eight colours straight from Okabe–Ito. Actual vs Budget is deep slate blue against amber, differing in hue *and* lightness; favourable vs unfavourable uses the Okabe–Ito green and vermillion. |
| Colour literals | 12 border literals → palette slot 0; 6 conditional-format values → palette green/vermillion; 4 Budget-series fills → palette slot 1; waterfall sentiment colours → palette; 14 redundant white background and border overrides deleted so the theme owns them. Slicer containers moved to the pale palette tint, fixing dark-text-on-mid-tone contrast. |
| Dynamic commentary | The three hard-coded narrative text boxes replaced by card visuals bound to the narrative measures, so the text re-evaluates on every slicer move. |
| EBIT card | The Budget vs Actuals KPI card rebound from the net profit measures to `Actual EBIT`, `Budget EBIT`, `EBIT Variance`, `EBIT Variance %` and `Actual EBIT Margin %`, with favourable/unfavourable conditional colouring. |
| OPEX caveat | Stated on the Budget vs Actuals page beside the variance sign convention, repeated inside the generated commentary, and carried in the descriptions of `Budget OPEX`, `Budget EBIT`, `EBIT Variance` and `EBIT Variance %`. **The source data was not altered.** |
| Titles | Explicit titles added to the two untitled charts. Two titles that were floating text boxes moved into the visuals' own title property and the text boxes deleted, so they scale, move and export with the visual. |
| Margin page fix | The "Gross Profit: Actual vs Budget by Account" chart was worse than the baseline review recorded. With `Actual Revenue` filtered to revenue accounts and the axis on a cost account, the intersection was empty, so every COGS bar was plotting *negative COGS* under a gross-profit title. It is now "COGS: Actual vs Budget by Account", on `Dim_Account[AccountName]` with a visual-level filter to the COGS section. |
| New page | **P&L Statement** — a matrix on `Dim_Account[AccountName]` in `PLSort` order showing actual, budget, favourable-signed variance and variance %, with Gross Profit and EBIT as calculated rows. The Interest and Net Profit accounts return blank on all four measures, so the matrix drops them without needing a filter, and the grand total is deliberately blank because a P&L does not add its own sections together. |
| Accessibility | Alt text added to every data visual across all four pages. There was none before. |
| Layout | A later pass corrected four visuals that were clipping their own content: the region × category matrix and the variance summary table were showing only part of their rows, the COGS by account chart only part of its accounts, and the P&L table filling half its width. Enlarging the COGS chart cost the COGS ratio chart above it 10px, which dropped it below the height needed to lay out four category bands, so the lowest-sorted category paginated off the bottom — restored to 222px. The narrative cards ignore `wordWrap`, so instead the three narrative measures were shortened to a single line: the longest text each can produce was checked across all 50,625 slicer states (every non-empty subset of Quarter × Region × Category × Channel), the widest being 768px inside a 950px usable card. The cards then dropped from 78px to 46px and the reclaimed space went to the charts. |

### 4.3 Repository

- Artefacts named `Financial_Performance_PnL.pbip` / `.Report` / `.SemanticModel`, with the paths
  inside `.pbip` and `definition.pbir` updated. Both `.platform` files keep the readable
  `displayName` **"Financial Performance (P&L)"**.
- `.gitignore` covers Power BI local state (`**/.pbi/localSettings.json`, `**/.pbi/cache.abf`),
  packaged binaries (`*.pbix`, `*.pbit`), Excel lock files and Windows shell cruft. The
  `localSettings.json` files hold DPAPI-encrypted, machine-bound security bindings and are never
  committed.
- `.gitattributes` normalises text to LF in the repository while Power BI keeps writing CRLF
  locally, and marks the workbook and other binaries as binary. **No file carries a UTF-8 BOM.**
- `README.md` covers the business question, what each page answers, the key DAX and model
  techniques, the screenshot gallery, the local run procedure, the variance convention, the OPEX
  caveat and the palette rationale.

---

## 5. Verification after the changes

Independent recomputation from the source workbook agrees with the live model to the cent on every
figure in §3.

| | §3 above | Live model after |
|---|---|---|
| Actual Revenue | 91,077.47 | **91,077.47** |
| Budget Revenue | 93,145.92 | **93,145.92** |
| Revenue Variance % | −2.221% | **−2.2207%** |
| Actual / Budget COGS Ratio | 42.19% / 40.85% | **42.192% / 40.851%** |
| Actual / Budget Gross Profit | 52,649.96 / 55,095.03 | **52,649.96 / 55,095.03** |
| Gross Margin % actual / budget | 57.808% / 59.149% | **57.808% / 59.149%** |
| Actual / Budget OPEX | 5,693.92 / 27,197.60 | **5,693.92 / 27,197.60** |
| EBIT actual / budget | 46,956.04 / 27,897.43 | **46,956.04 / 27,897.43** |
| EBIT Variance % | +68.317% | **+68.317%** |

All four brand categories and all four regions reconcile exactly to §3.2 and §3.3. The P&L statement
is internally consistent: the three revenue lines sum to Actual Revenue, the four COGS lines to
Actual COGS, the five OPEX lines to Actual OPEX, and the calculated Gross Profit and EBIT rows equal
the corresponding measures.

**The display bugs, confirmed fixed against the engine:**

| Measure | Before | After |
|---|---|---|
| `Gross Margin vs Budget Label` | `▼ 0.01pp` | **`▼ 1.34pp`** |
| `Revenue vs Budget Label` | `▼ -2.2%` (double negative) | **`▼ 2.2%`** |
| `Actual COGS Ratio` | `0.4219` (unformatted) | **`42.2%`** |

**Time intelligence**, impossible before the date table was fixed, now resolves: `TOTALYTD`
accumulates 18,584.67 → 40,064.35 → 64,293.00 → 91,077.47 across the four quarters, and
`DATEADD(Dim_Date[Date], -1, QUARTER)` returns the prior quarter. `SAMEPERIODLASTYEAR` correctly
returns blank, since the dataset holds 2024 only.

**Narrative measures** were tested across six filter states. At the default state they reproduce the
figures in §3 exactly. Testing the filtered states exposed three wording defects, all since fixed:

| Defect found | Fix |
|---|---|
| "Review the Skincare COGS **overrun** (0.0pp **below** budget)" — the noun was hard-coded while the direction was dynamic | each clause now branches on the sign of its own figure, and reports "no category is running above its budgeted COGS ratio" when none is |
| "Direct is the largest revenue **drag** at €42" — €42 is favourable; with every channel above budget there is no drag | reports "every channel is at or above budget, X the weakest at …" instead |
| "Europe is the **strongest region** at −2.5%" — vacuous when one region is selected | collapses to "Europe is −2.5% against budget" when a single region is in context, mirroring the single-category branch |

An empty selection returns "No revenue in the current selection." rather than an error.

**Best Practice Analyzer:** **207 → 33 violations.** All 33 remaining are deliberate:

| Count | Rule | Why it stays |
|---|---|---|
| 24 | Whole numbers should be formatted with thousands separators and no decimals | The rule only accepts a **dollar** sign in the format string. The 18 euro-denominated measures, the 5 text measures and the one datetime measure cannot satisfy it without dropping the currency symbol. The rule encodes a USD house style; this model reports in EUR. |
| 5 | Provide format string for measures | The two label measures and three narrative measures return **text**. The rule does not exclude string-typed measures, for which a numeric format string is meaningless. |
| 2 | Ensure tables have relationships | `_Measures Table` (a measure holder by design) and `Refresh Stamp` (a single-row utility table). Both are correct as they are. |
| 1 | Hide foreign keys | Flags `Dim_Account[AccountName]`. The rule matches on column **name only**, ignoring the table, so it sees the fact's same-named foreign key. This is the dimension's primary key and the P&L row axis, so it must stay visible. |
| 1 | Minimize Power Query transformations | The `Dim_Account` query's added `PLSort` step. Five steps on a 16-row dimension. |

## 6. Deliberate exceptions

- **Dimension attributes were kept, not deleted.** `Dim_Brand[BrandName, LaunchYear, PriceSegment]`,
  `Dim_Region[Currency, Cluster, MarketStage]` and `Dim_Channel[ChannelType, CostLevel]` are unused
  by any visual, but a star-schema dimension is expected to carry descriptive attributes and a
  reader may well slice by them. Deleting them would narrow the model's apparent scope for no real
  gain. The genuinely redundant columns were removed.
- **The calendar stays daily**, at 366 rows against a monthly fact, because date-table marking and
  time intelligence require a contiguous daily date column.
- **No unit is asserted.** The source data is EUR-denominated, so `€` is used throughout, but
  nothing in it states a scale (thousands, millions), so none is claimed. The README and the P&L
  page both say the amounts are shown exactly as supplied.
- **The source data was not altered**, including the flat OPEX budget allocation.

## 7. Still open

- **`SourceFolder` ships with a placeholder value**, so no machine's directory layout or user name
  enters the repository. It must be pointed at the clone's `data` folder once before the first
  refresh; the README says so.
- **The published report predates the layout pass** in §4.2 and will keep serving the earlier
  version until it is republished.
- **The waterfall's two smallest data labels sit close enough to touch** at the current width. Not
  corrected, because the values they label are legible and moving them risks the layout budget the
  rest of the page now depends on.
