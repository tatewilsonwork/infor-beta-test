# SEDI extraction reference

How to obtain and parse a SEDI "Insider Information by Issuer" report for the ownership workbook.

## Obtaining the report (analyst, manual)

SEDI (`https://www.sedi.ca/sedi/`) is the System for Electronic Disclosure by Insiders, operated by
the Canadian Securities Administrators. It is **Canadian-only** and **cannot be auto-fetched**: the
"Access Public Filings" path is behind a Radware/ShieldSquare bot wall that returns a CAPTCHA to any
scripted client, and CAPTCHAs are not solved by automation. So the analyst downloads the report by
hand and attaches the PDF:

> English → **Access Public Filings** → **View Summary Reports** → **Insider information by issuer** →
> type the issuer name in the search bar → **View** → **Download PDF Report**

The resulting PDF (URL `SVTIIBIviewResults`) is fully text-extractable — no OCR needed.

There is **no US equivalent** wired into this skill: US insider holdings live in SEC Forms 3/4/5 on
EDGAR. The ownership slide therefore applies only to **Canadian reporting issuers**.

## What the report contains (and does not)

It lists **insiders** only — not the full shareholder base. Relationship codes:

| Code | Relationship to issuer |
|------|------------------------|
| 1 | Issuer |
| 2 | Subsidiary of issuer |
| 3 | 10% security holder of issuer |
| 4 | Director of issuer |
| 5 | Senior officer of issuer |
| 6 | Director/senior officer of a 10% holder |
| 7 | Director/senior officer of an insider or subsidiary (other than 4/5/6) |
| 8 | Deemed insider — 6 months before becoming an insider |

There is **no sub-10% institutional or retail float** here — that is the Bloomberg-sourced
institutional side of the slide, filled separately. The report is also **cumulative since ~2003**:
it includes long-departed insiders, which is why the current-insider filter (below) is essential.
Balances are share/unit **counts only** — no prices.

## Reading an insider block

Each insider is a block:

```
Insider Name: Barrenechea, Mark James   Insider Relationship: 4 - Director of Issuer, 5 - Senior Officer of Issuer
Ceased to be Insider: 2025-08-11
  2025-03-31 Common Shares OTEX Common 1,219,092
  2024-08-19 Options All OTEX Option Plans(Common Shares OTEX Common) 3,666,720  3,666,720
  2025-06-20 Performance Share Units (Common Shares) 435,928  435,928
  2025-06-20 Restricted Share Units (Common Shares) 217,969  217,969
```

Each security row is: `<date> <security designation> <registered holder> <closing balance> [<calculated / underlying>]`.

### Rules

1. **Current insiders only.** Keep the block iff `Ceased to be Insider:` is `Not Applicable`. Drop any
   block with a ceased date (the Barrenechea example above is *ceased* — it would be dropped).
2. **Common shares only.** Sum only rows whose security designation is **`Common Shares`**. Exclude
   `Options …`, `Deferred Share Units …`, `Performance Share Units …`, `Restricted Share Units …`,
   warrants, and convertibles. (Those are equity awards, not common shares.)
3. **Sum tranches as a list.** One insider can hold common shares under several registered holders,
   e.g. Randy Fowlie:
   ```
   2025-12-01 Common Shares OTEX Common 193,000
   2025-04-14 Common Shares OTEX Common  Armour Stone Capital Ltd. 0
   2020-06-08 Common Shares OTEX Common  RCF Advisors Ltd. 0
   2008-11-13 Common Shares OTEX Common  RRSP 0
   2008-11-13 Common Shares OTEX Common  Spousal RRSP 0
   ```
   Pass `[193000, 0, 0, 0, 0]` so the workbook writes `=193000+0+0+0+0` (auditable, not hand-summed).
4. **Most recent date** = the latest date among that insider's common-share rows.
5. **Adjusted name + role.** `Last, First Middle` → `First Last`; append the role from the
   relationship code plus the specific title (CEO / CFO / Chairman / …) confirmed from the company
   website or LinkedIn. Prioritise the largest common-share holders (only the top 12 appear on the
   slide).

## Worked extraction (current insiders)

| SEDI name | Relationship | Common-share tranches | Most recent | Adjusted name |
|-----------|-------------|----------------------|-------------|----------------|
| Fowlie, Randy | 4 – Director | `[193000,0,0,0,0]` | 2025-12-01 | Randy Fowlie (Director) |
| Acedo, Michael Fernando | 5 – Senior Officer | `[13371]` | 2025-11-10 | Michael Acedo (Chief Legal Officer) |
| Antoun, Ayman | 5 – Senior Officer | `[0]` | 2026-04-20 | Ayman Antoun (President) |

Pass each as an `InsiderHolding` to `build_ownership_workbook` (see SKILL.md Step 9).
