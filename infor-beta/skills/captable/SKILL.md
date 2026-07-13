---
name: captable
description: >
  Use this skill when extracting financial data from MD&A, 10-K, 10-Q, annual reports, or financial
  statements to populate a capitalization table. Activates on /captable and on tasks involving
  shares outstanding, debt schedules, lease obligations, options/RSU/warrant tables, convertible
  debentures, cash balances, preferred shares, or non-controlling interest sourced from company filings.
version: 0.5.26
allowed-tools: [Read, Bash, Write, Glob, WebSearch, WebFetch]
---

# INFOR Capitalization Table — Workflow & Domain Knowledge

This skill guides you through populating the INFOR capitalization table template from a CapIQ ticker and attached financial statements, and provides domain knowledge for accurate data extraction.

Today's date is available from the system context (`currentDate`) — do not shell out to `date`. Template location and working directory are resolved inline in Step 2.

## Conductor-mode handoff (read first when running under the conductor)

When invoked as a stage of a conductor plan, the environment carries `$STAGE_INPUTS`, `$STAGE_OUTPUTS`, and `$DEAL_DIR`:

- Read your inputs (ticker, optional company facts) from `$STAGE_INPUTS` instead of asking the analyst. If `$STAGE_INPUTS` is missing fields you need, surface the gap by writing `{"error": "missing input: <field>"}` to `$STAGE_OUTPUTS` and stop — the conductor halts on that.
- `$STAGE_INPUTS` may also carry `ltm_revenue` and `ltm_adj_ebitda` (millions, in the **filing's reporting currency**) handed off by the upstream `ltm-metrics` stage. Use them in Step 6b to populate the cap table's LTM valuation column (D47 / D48). They are optional — when absent (e.g. direct `/captable` invocation, or a plan with no `ltm-metrics` stage), Step 6b restores the CapIQ fallback formulas instead.
- Write the cap table to `$DEAL_DIR/artefacts/<SANITIZED_TICKER> - Capitalization Table.xlsx` (NOT cwd). Bootstrap `$DEAL_DIR/artefacts/` if it doesn't exist.
- At the end of Step 8, also write the structured handoff:
  ```bash
  python -c "import json,os; json.dump({'workbook_path': os.environ['OUTPUT']}, open(os.environ['STAGE_OUTPUTS'], 'w'))"
  ```
  The conductor will not proceed past this stage until `$STAGE_OUTPUTS` exists and parses as JSON.

When `$STAGE_OUTPUTS` is **unset** (direct `/captable` invocation), follow the workflow below as-is — output lands in cwd, no JSON handoff needed.

---

## Workflow Steps

### Step 1 — Collect Inputs

Ask the user for the CapIQ ticker if not already provided, and confirm at least one document is attached (10-K, 10-Q, annual report, or financial statement PDF).

If either is missing, ask in a single message:

> "Please provide the CapIQ ticker for the company and attach the relevant financial statements or MD&A with your reply.
>
> - **Ticker format:** Exchange:Ticker — for example `NasdaqGS:MSFT`
> - **Documents to attach:** Most recent 10-K or annual report — I need the balance sheet, long-term debt note, lease footnote, equity compensation footnote, and shares outstanding disclosure to fill in the template."

Wait for both the ticker and at least one attached document before proceeding.

> **Build the cap table off the *most recent* reported statement only.** If the analyst has attached more than one period (e.g. a prior full fiscal year alongside the latest 10-Q), the older statement is there for the `ltm-metrics` LTM bridge, **not** for the cap table. Always source shares outstanding, debt, leases, cash, etc. from the latest reported balance sheet — never from the older FY statement.

---

### Step 2 — Locate and Copy the Template

Resolve the template via the plugin's shared helper:
```bash
bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/find_template.sh" "INFOR Cap Table Template.xlsx"
```

Sanitize the ticker for use as a filename via the shared helper (e.g., `NasdaqGS:MSFT` → `NasdaqGS-MSFT`):
```bash
SANITIZED_TICKER=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/sanitize_name.sh" "$TICKER")
```

Copy the template to the current working directory using this exact shell pattern (note the quoting — required because the path contains spaces):
```bash
TEMPLATE=$(bash "${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/find_template.sh" "INFOR Cap Table Template.xlsx")
OUTPUT="./$SANITIZED_TICKER - Capitalization Table.xlsx"
cp "$TEMPLATE" "$OUTPUT" && echo "COPY_OK" || echo "COPY_FAILED"
```

Then verify the output file exists and has non-zero size:
```bash
ls -lh "$OUTPUT"
```

**If the result is `COPY_FAILED`, the file is missing, or its size is 0 bytes — STOP immediately. Do NOT proceed. Do NOT create a manual cap table or output data in any other format. Tell the user:**

> "I could not copy the INFOR Cap Table Template. Please confirm the file `INFOR Cap Table Template.xlsx` exists in the `templates/` folder at the root of the `infor-beta` plugin repository."

---

### Step 3 — Update Header Inputs (openpyxl)

Open the copied file with openpyxl (NOT data_only — preserve all formulas).

Update this cell in the `Cap with Links` sheet:

| Cell | Value |
|------|-------|
| F3 | CapIQ ticker exactly as provided (e.g., `NasdaqGS:MSFT`) |

**Do not modify these header cells:** F4, F5, F10, F11, or any formula cell. Cell F6 already contains a `=TODAY()-1` formula — leave it untouched.

**F7 (FX Rate) and F16 (Share Price) are intentionally empty in the template and you fill them from the web (see Step 3b).** Each carries a cell **comment** holding the CapIQ formula the analyst pastes back to refresh the value live — **never delete or overwrite that comment.** Write only the cell value; openpyxl leaves an existing comment intact.

---

### Step 3b — Populate FX Rate (F7) and Share Price (F16) from the web

The template no longer carries live CapIQ formulas in F7 and F16; you supply a current static value from the web so the cap table computes immediately, and the analyst can later paste the commented CapIQ formula to refresh it. Source both from a reputable market data site, as of the F6 date (yesterday — the latest prior trading close).

- **F16 — Share Price.** The price of one common share, expressed in the **Output currency (F5)**. Read F5 first (default `CAD`). If the share trades in a different currency, convert the quoted price into F5 using the same FX rate you write to F7. Write a numeric value (e.g. `42.18`), not a string.
- **F7 — FX Rate.** Defined by the template math as **Output-currency units per 1 unit of the filing's reporting (Input) currency**: downstream cells convert filing-currency figures to the output currency by *multiplying* by F7 (e.g. `F24 = F122*F7`), and Section II/III strike comparisons assume `strike_output = strike_filing * F7`. So if the filing reports in USD and F5 is `CAD`, F7 is **CAD per USD** (e.g. `1.37`). If the filing currency equals F5, write `1.0`.

Determine the filing's reporting currency from the attached documents. Write both values with the Palatino-9 blue input font `Font(name="Palatino Linotype", size=9, color="0000FF")` (they are hardcoded — a bare `Font(color=...)` would reset them to Calibri 11), and leave their CapIQ-formula comments untouched. Note the source and as-of date for each in the Step 8 summary (not as a cell comment — the comment slot is reserved for the CapIQ formula).

---

### Step 4 — Extract Financial Data

Read the attached documents and extract all items listed in the Domain Reference section below. Use the most recent balance sheet date available.

**Garbled PDFs:** if a filing's text comes through scrambled (CID-font glyph soup, U+FFFD replacement characters, or blank pages), do **not** transcribe the garbage. Use the shared `${CLAUDE_PLUGIN_ROOT:-./infor-beta}/scripts/pdf_extract.py` helper, which detects garble and falls back to rendering the pages + tesseract OCR: `from pdf_extract import extract_pdf_text; r = extract_pdf_text(path)` — read `r.text`, and if `r.garbled` is still true (tesseract unavailable) tell the analyst rather than guessing the numbers.

**Source tracking:** For every value extracted, record its source in this format:
`"[Document Name] - Page [#], [Section Name]"`
You will attach these as cell comments in Step 6.

---

### Step 5 — Subsequent Events Scan

Capital structures can change materially between the balance sheet date and today. Run both checks below and apply adjustments **on top of the Step 4 figures** so the cap table reflects the current capital structure, not a stale snapshot.

**5a — Filing subsequent-events note.** Locate the subsequent-events section in the attached filing:
- US GAAP: "Subsequent Events" note (ASC 855), typically the final footnote.
- IFRS: "Events After the Reporting Period" note (IAS 10).
- Also scan the MD&A "Recent Developments" / "Outlook" sections — material capital events are often disclosed there even if the formal sub-events note is silent.

For every event that changes the cap table (see the Cap-Table-Impacting Events list in the Domain Reference), adjust the relevant row in Step 4 before writing. Do not create a separate "subsequent events" section — fold the adjustments into the existing rows and note the event in the row's cell comment.

**Row labeling — REQUIRED.** Whenever a subsequent event adds a **new row** to the cap table (e.g., a new debt tranche, new share issuance, new convertible), the col B label MUST be prefixed `"Sub-Event: "` followed by a brief description. Examples:
- `"Sub-Event: $500M 5.25% Senior Notes due 2034"` (new row in Section IV)
- `"Sub-Event: Feb 2026 follow-on offering"` (new row in Section VII)
- `"Sub-Event: 2026 Convertible Notes"` (new row in Section III, with matching `"Sub-Event: 2026 Convertible Notes (if OTM)"` row in Section IV)

When a subsequent event **modifies an existing row** (e.g., buyback reducing share count, revolver draw updating an existing facility balance) rather than adding a new one, leave the col B label unchanged — the event is disclosed only in the cell comment per the source-tracking format below.

**5b — Newsroom web screen.** Search **only company-issued sources** — the company's investor-relations newsroom and its own press releases — for items dated **after the balance sheet date** through today. Do **NOT** use third-party news sites, analyst reports, aggregators, or wire services (Reuters, Bloomberg, Yahoo Finance, Seeking Alpha, etc.) as the primary source. A press release hosted on a wire (e.g., Business Wire, GlobeNewswire, PR Newswire, CNW) is acceptable only when it is the company's own release republished by the wire — confirm the release is issued by the company, not written about the company.

Use WebSearch with queries scoped to the company's own domain, for example:
- `"[Company Name]" (offering OR issuance OR redemption OR buyback OR "share repurchase" OR acquisition OR dividend OR "credit facility" OR convertible) site:[IR domain]`
- `"[Company Name]" press release [Month Year] site:[company domain]` for each month between the balance sheet date and today

Then WebFetch the IR newsroom index page directly to confirm no material company releases were missed. If you cannot locate an IR/newsroom page on the company's domain, note this in the Step 8 summary and skip 5b rather than substituting a third-party source.

**Deduplication rule:** Before applying a newsroom item, confirm it is NOT already captured in 5a (the filing's sub-events note) or in any footnote read during Step 4. Common duplicates:
- Post-year-end debt refinancings already disclosed in the long-term debt note
- Share issuances announced pre-filing and cross-referenced in the capital stock note
- Dividends declared before the filing date and mentioned in MD&A

If unsure whether an event is already captured, prefer the filing's disclosure and skip the newsroom version — do not double-count.

**Source tracking for sub-events:** Use this comment format, appended to the existing source citation on the affected row:
`"[Original source] + Subsequent event [YYYY-MM-DD]: [brief description] per [filing note / press release URL]"`

Example:
`"Rogers 2024 Annual Report - Page 87, Note 12: Long-Term Debt + Subsequent event 2026-02-10: $500M senior notes issued per Rogers press release https://..."`

If no subsequent events are found in either scan, note this in the Step 8 summary — do not invent events.

---

### Step 6 — Write Extracted Data to Excel

**CRITICAL: All extracted data must be written into the copied INFOR template .xlsx file using openpyxl. Never output the data as a markdown table, plain text, or any format other than the .xlsx file. If openpyxl is not available, stop and tell the user: "Please install openpyxl: `pip install openpyxl`"**

Using openpyxl (write mode, preserve formulas — do NOT use data_only=True):

Write each section into the correct rows and columns per the Template Row Map below. Use Python `datetime.date` objects for date cells and numeric values (not strings) for amount cells.

**Color coding:** Mark every hardcoded value cell you write with the blue input font `Font(name="Palatino Linotype", size=9, color="0000FF")`. Do NOT use a bare `Font(color="0000FF")` — it carries no typeface, so it resets the cell to the Calibri 11 default and drops the template's Palatino 9. Define it once and reuse it:
```python
from openpyxl.styles import Font
BLUE = Font(name="Palatino Linotype", size=9, color="0000FF")  # hardcoded-input font
# usage: ws["F16"].value = 21.20; ws["F16"].font = BLUE
```
Do NOT recolor formula cells.

**Cell comments:** Attach a single openpyxl `Comment` to the **first cell in each row** (col B). One comment per row only:
```python
from openpyxl.comments import Comment
ws["B108"].comment = Comment("Rogers 2024 Annual Report - Page 87, Note 12: Long-Term Debt", "INFOR")
```

**Never write to formula total cells** E82, F82, F84, F85, F86, F104, F122, F143, F164, F186.

Write the file.

---

### Step 6b — Populate LTM Valuation Metrics (D47 / D48)

The Financial Metrics block reads its **LTM** column from `D47` (Revenue) and `D48` (Adj. EBITDA): `D34=D47`, `D35=D48`, which in turn drive `EV / Revenue` (D39), `EV / Adj. EBITDA` (D40) and the Adj. EBITDA margin (D36). The template ships with `D47`/`D48` **empty** — you fill them here.

**Currency / units.** The CapIQ estimate columns next to them (`E47`/`F47`, `E48`/`F48`) are expressed in the **Output currency (F5), in millions**. Whatever you write to `D47`/`D48` must match: millions, in F5 currency.

**Case A — LTM values supplied (`ltm_revenue` / `ltm_adj_ebitda` in `$STAGE_INPUTS`).** These arrive in the filing's reporting currency, in millions. Convert to the Output currency the same way the rest of the template does — by **multiplying by F7** (Output per Input/filing currency; F7 is `1.0` when the filing currency equals F5, so this is a safe no-op then). Write each as a formula embedding the figure so the conversion stays auditable, and apply blue font (it is an authored hardcoded-derived value, like the Section IV OTM rows):

```python
if ltm_revenue is not None:
    c = ws["D47"]; c.value = f"={ltm_revenue}*F7"; c.font = BLUE  # Palatino-9 blue (see Color coding)
if ltm_adj_ebitda is not None:
    c = ws["D48"]; c.value = f"={ltm_adj_ebitda}*F7"; c.font = BLUE
```

**Case B — no LTM values (direct invocation / no `ltm-metrics` stage).** Restore the CapIQ LTM formulas so the cap table still auto-populates in Excel with the CapIQ add-in. Do **not** color these blue (they are formula cells):

```python
ws["D47"] = '=_xll.SNL.Clients.Office.Excel.Functions.SPG($F$3,"IQ_REV",D$33,$F$6,"Options:Mag=Millions,NA=NA,Curr="&$F$5)'
ws["D48"] = '=_xll.SNL.Clients.Office.Excel.Functions.SPG($F$3,"SP_EBITDA",D$33,$F$6,"Options:Mag=Millions,NA=NA,Curr="&$F$5)'
```

Write the file.

---

### Step 7 — Recalculate and Verify

Run the recalc script if available:
```bash
python "[xlsx_skill_scripts_path]/recalc.py" "./[SANITIZED_TICKER] - Capitalization Table.xlsx"
```

If unavailable, skip this step. Then check for `#REF!`, `#DIV/0!`, `#VALUE!`, or `#NAME?` errors in cells you wrote and fix any found.

Perform the cross-checks listed in the Domain Reference below before delivering.

---

### Step 8 — Summary

Report to the user:

1. **Output file:** path to the saved file
2. **Web-sourced market values:** the FX rate (F7) and share price (F16) you wrote, each with its source and as-of date. Remind the user these are static snapshots — the analyst can paste the CapIQ formula stored in each cell's comment to refresh them live in Excel.
3. **LTM valuation metrics (D47 / D48):** state whether LTM Revenue and Adj. EBITDA came from the `ltm-metrics` handoff (give the figures and note `*F7` FX conversion to the F5 currency) or fell back to the CapIQ `IQ_REV`/`SP_EBITDA` formula because no LTM values were supplied.
4. **CapIQ auto-populated fields** (will refresh when file is opened in Excel with CapIQ connection): company name, FYE, currency, report dates, revenue & EBITDA consensus estimates, analyst coverage, average target price
5. **Fields populated from MD&A:** list each section and what was found
6. **Subsequent events applied:** for each adjustment made in Step 5, list the event date, description, source (filing note vs. press release), and which row it adjusted. If none were found, state "No subsequent events identified in filing sub-events note or newsroom screen through [today's date]."
7. **Items not found:** any line items missing from the attached documents (user should fill these manually in blue)
8. **Reminder:** Open in Excel with the CapIQ add-in active to refresh market data and the commented F7/F16 formulas

---

## Domain Reference

### Template Row Map

| Section | Description | Input Columns | Rows | Formula Totals (do NOT overwrite) |
|---------|-------------|---------------|------|----------------------------------|
| Header | Ticker | F3 | — | — |
| Header | FX Rate (web-sourced, Output per Input), Share Price (web-sourced, in F5 currency) | F7, F16 | — | — |
| LTM | LTM Revenue (D47), LTM Adj. EBITDA (D48) — `=<value>*F7` from `ltm-metrics`, else CapIQ fallback (Step 6b) | D47, D48 | — | D34/D35 read these (do NOT overwrite D34/D35) |
| I | Preferred Shares, NCI | F52, F53 | — | — |
| II | Options / Warrants / RSUs / DSUs | B (type), C (amount, M shares), D (strike) | 57–81 | E82, F82, F84, F85, F86 |
| III | Convertible Debentures / Preferred | B (type), C (face, $M), D (shares/1000), E (strike) | 90–103 | F104 |
| IV | Debt Schedule | B (facility), E (as-of date — financial statement date, NOT maturity date), F (amount, $M) | 108–121 | F122 |
| V | Lease Obligations | B (type), E (date), F (amount, $M) | 126–142 | F143 |
| VI | Cash & Equivalents | B (type), E (date), F (amount, $M) | 147–163 | F164 |
| VII | Basic Shares Outstanding | B (description), E (date), F (amount, M shares) | 168–185 | F186 |

**Never write to formula total cells** E82, F82, F84, F85, F86, F104, F122, F143, F164, F186. These are SUM/formula cells that auto-total the rows above them.

### Unit Conventions

- **Always match the template's existing unit.** The template uses millions for all figures.
  - Filing reports in **thousands** → divide by 1,000
  - Filing reports in **full units** (e.g., 22,841,361 shares) → divide by 1,000,000
  - Filing already in **millions** → enter as-is
- **Currency:** Never convert — input exactly as stated in the filing.
- **Dates:** Use the balance sheet date for debt, cash, and lease figures. For shares (Section VII col E), use the subsequent-event date from the capital stock note.

### Cell Comments — Source Citations

Every hardcoded cell must have an openpyxl `Comment` attached (one per row, on col B):
```
"[Document Name] - Page [#], [Section Name]"
```
Examples:
- `"Rogers 2024 Annual Report - Page 87, Note 12: Long-Term Debt"`
- `"MSFT 2024 10-K - Page 14, Consolidated Balance Sheet"`

### GAAP vs. IFRS Differences

| Item | US GAAP | IFRS |
|------|---------|------|
| Debt note label | "Long-Term Debt" | "Borrowings" or "Loans and Borrowings" |
| Lease standard | ASC 842 — finance + operating leases separately | IFRS 16 — single lease liability, row 126 only |
| Equity comp note | "Stock-Based Compensation" | "Share-Based Payments" |

**IFRS lease:** Write F126 as a formula `ws["F126"] = current + noncurrent` — not a scalar. Leave row 127 blank.

### Where to Find Each Data Item

**Basic Shares Outstanding:** Capital stock note — use the subsequent-event date (e.g., "As at Feb 18, 2026, 22,841,361 shares outstanding"). Divide full-unit counts by 1,000,000.

**Cash:** Balance sheet first line of current assets. Also capture short-term investments and restricted cash separately.

**Long-Term Debt:** "Long-Term Debt" or "Borrowings" footnote. Enter each tranche at face value; add a separate negative row for unamortized discount/issuance costs. Always include the revolving credit facility even at $0 — label "Revolving Credit Facility (undrawn)".

**Leases:** ASC 842 or IFRS 16 footnote. Use the discounted lease liability balance, not undiscounted payments.

**Options/Warrants/RSUs/DSUs:** Stock-based compensation footnote. Enter one row per exercise-price tranche for options — do NOT aggregate to WAEP. RSUs/DSUs use $0 strike. **Always exclude PSUs, and exclude any RSUs/DSUs that are exclusively cash-settled** (they pay out in cash rather than shares, so they don't factor into ITM dilutive share count). Check the equity comp footnote for settlement terms — language like "settled in cash," "cash-settled only," or "no share issuance" signals exclusion. If an RSU/DSU plan allows share OR cash settlement at the company's discretion, include it.

**Options strike — FX conversion required whenever filing currency ≠ Output currency (F5).** The template's ITM calculation compares the strike in col D against the share price in F$16, which is stated in the Output currency (F5). Unlike Section III (where the ITM formula is authored per-row and can apply FX inline), Section II strikes are compared natively by template formulas you cannot modify — so the strike **value written to col D must already be in the Output currency**. F7 is **Output-currency per Input(filing)-currency** (you set it in Step 3b), so convert by **multiplying**: `strike_output = strike_filing * F7`.

In the cell comment on col B, note both the original filing-currency strike and the FX rate applied, e.g., `"Rogers 2024 Annual Report - Page 102, Note 20: Stock Options. Original strike CAD $45.20 converted to USD at F7 = 0.74"`. Apply blue font to col D as with any hardcoded value.

**Convertible Debentures / Convertible Preferreds:** Face amount (col C), shares per $1,000 face = 1,000 / conversion price (col D), conversion price (col E). **Also add a matching row in Section IV (Debt)** so the face is captured as debt only when the convert is out-of-the-money:
- Label col B: `"[Convert name] (if OTM)"`
- Date col E: the as-of date of the financial statement the convert face amount is sourced from (NOT the maturity of the convert — col E is an as-of date throughout Section IV).
- Amount col F: enter as an **IF formula** referencing the Section III row. If strike < share price (ITM) the convert will convert to shares and contributes $0 to debt; otherwise the face (C[row]) flows into total debt.
  - **FX-aware comparison — required whenever filing currency ≠ Output currency (F5):** share price (F$16) is stated in the Output currency (F5); strike (col E in Section III) is in the filing's reporting currency. F7 holds the FX rate (Output per Input — set in Step 3b). You MUST align the two sides before comparing, or the ITM/OTM classification will be wrong whenever F5 and the filing currency differ. Multiply the strike by F$7 to bring it into F5 currency. Same-currency example (row 90): `=IF(E90<F$16,0,C90)`. Cross-currency example (row 90) converting the strike into F5 currency: `=IF(E90*F$7<F$16,0,C90)`.
- Apply blue font to col B, E, and F (the F cell holds a hardcoded formula you authored, not a template formula, so color it).

**Preferred Shares / NCI:** Balance sheet equity section. Enter 0 if none.

### Cap-Table-Impacting Subsequent Events

Apply the following event types when found in the filing's sub-events note or in post-filing press releases. Ignore events that don't move any row in the template (operational updates, guidance changes, leadership moves, customer wins).

| Event | Section/Row affected | Adjustment |
|-------|---------------------|------------|
| Equity issuance / follow-on / PIPE | VII — Basic Shares | Add issued shares to a new row with col E = closing date |
| Share buyback completed or ASR settled | VII — Basic Shares | Subtract repurchased shares; note buyback authorization amount and date in comment |
| Stock split / consolidation | VII — Basic Shares (and II, III strike prices) | Apply ratio to shares and to every option/convert strike |
| New debt issuance (notes, term loan, revolver upsize) | IV — Debt Schedule | Add a new row at face with col E = issuance date |
| Debt redemption / repayment / tender | IV — Debt Schedule | Reduce or remove the redeemed tranche |
| Revolver draw or paydown | IV — Debt Schedule | Update the revolver row balance and rename to "Revolving Credit Facility (drawn $XM)" if drawn |
| Convertible issuance | III (new row) **and** IV (matching OTM/ITM row) | Add face, shares/1000, strike to Section III; add linked `=IF(...,0,C[row])` row in IV |
| Convertible conversion / redemption | III + IV | Zero out or remove both the Section III and Section IV rows |
| Acquisition closing (cash-funded) | VI — Cash | Reduce cash by purchase price; add assumed debt to IV if applicable |
| Acquisition closing (stock-funded) | VII — Basic Shares | Add consideration shares issued |
| Dividend declared + paid after BS date | VI — Cash | Reduce cash by total dividend paid |
| Option grant / RSU grant / exercise post BS date | II — Options/RSUs | Add new tranche or reduce outstanding count if exercised |
| Lease signed or terminated (material) | V — Leases | Adjust lease liability balance |
| Preferred share issuance / redemption | I (F52) or III if convertible | Update balance or add Section III row |

**When in doubt, include and flag.** If a press release describes a material capital event that doesn't clearly map to a row, apply your best interpretation, note the uncertainty in the cell comment, and call it out in the Step 8 summary so the user can review.

### Fallback — No Documents Attached

- **US filers:** `https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=[ticker]&type=10-K`
- **Canadian filers:** `https://www.sedarplus.ca`

### Common Pitfalls

| Issue | Guidance |
|-------|----------|
| Face vs. carrying value | Enter face value; add negative row for issuance costs |
| Revolver at $0 | Always include — label "Revolving Credit Facility (undrawn)" |
| IFRS leases | Write F126 as formula, not scalar; leave row 127 blank |
| PSUs | Always exclude from Section II |
| Cash-settled RSUs/DSUs | Exclude from Section II — they pay in cash, not shares |
| Options aggregated | Enter one row per tranche, not single WAEP row |
| Options strike currency | Col D must be in Output currency (F5) — apply F7 to strike before writing whenever filing currency ≠ F5, otherwise the template's ITM comparison against F$16 breaks |
| Shares date | Use capital stock note subsequent-event date, not balance sheet date |
| RSU/DSU strike | Use $0 |
| Convertible preferred | If convertible → Section III; if not → F52 |
| Convertibles & debt | Every Section III row needs a matching Section IV row with `=IF(E[row]<F$16,0,C[row])` so OTM converts flow into debt — apply F$7 to align strike and share price whenever filing currency ≠ F5 |
| Debt col E | Column E in Section IV is the **as-of date** of the information (financial statement date), NOT the debt's maturity date |

### Cross-Checks

1. F17 (Basic Shares) = F186 — formula-linked
2. F122 (Total Debt) should tie to balance sheet carrying value
3. F164 (Total Cash) should be positive
4. Section II contains no PSUs and no cash-settled RSUs/DSUs
5. Revolver appears in Section IV even at $0
6. IFRS: only row 126 populated in Section V; row 127 blank
7. Options appear as one row per tranche, not single WAEP row
8. Section VII col E dates reflect capital stock note subsequent-event date
9. Every row populated in Section III has a matching `=IF(E[row]<F$16,0,C[row])` row in Section IV, with F$7 applied to the strike or share price whenever filing currency ≠ Output currency (F5)
10. Section II option/warrant strikes in col D are stated in the Output currency (F5) — if filing currency ≠ F5, each strike has been converted by multiplying by F7 (Output per Input) so the template's ITM comparison against F$16 works correctly
11. F7 (FX Rate) and F16 (Share Price) carry web-sourced numeric values in blue font, and their CapIQ-formula comments are still intact for the analyst to refresh
12. D47/D48 (LTM Revenue / Adj. EBITDA) are populated — either `=<value>*F7` in blue from the `ltm-metrics` handoff (millions, FX-converted to F5 currency), or the CapIQ `IQ_REV`/`SP_EBITDA` fallback when no LTM values were supplied — so D34/D35 and the EV multiples (D39/D40) resolve to numbers, not `n/a`
13. Subsequent-events scan completed: both the filing's sub-events note (ASC 855 / IAS 10) and a company-source-only newsroom screen (IR page and company press releases — no third-party news wires or analyst coverage) have been reviewed. Every applied adjustment is deduplicated against the filing's disclosures and documented in the relevant row's cell comment
