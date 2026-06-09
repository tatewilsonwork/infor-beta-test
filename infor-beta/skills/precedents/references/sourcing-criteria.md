# Precedent Transactions — Sourcing Criteria

Detailed rules behind the `precedents` SKILL.md. Carries over the production
`infor-workflows` `precedents-infor` sourcing criteria, adapted to the multi-metric
template: the metric family is now a choice, a disclosed **multiple is preferred** and is
written straight into the ratio column, and the multi-filing LTM stub calc has been dropped
(use the most recent reported figure as the LTM/NTM proxy instead).

## Choosing the metric family

Pick **one** family for the whole table, by the target's industry.

**Operating** (the default — software, industrials, consumer, healthcare, business services,
most non-financials): fill **Revenue** (LTM/NTM) and **Adj. EBITDA** (LTM/NTM); the table
shows **EV/Revenue** and **EV/EBITDA**.

**Financial** (banks, thrifts, insurers, asset & wealth managers, broker-dealers, specialty
lenders — businesses valued off the balance sheet / earnings rather than EBITDA): fill **Net
Income** (LTM/NTM), **Book Value**, and **Tangible Book Value**; the table shows **P/E**,
**P/B**, and **P/TBV**. EBITDA is not meaningful for these, so leave the operating columns
blank.

Edge cases: a fintech or exchange that trades like software → operating. A balance-sheet-heavy
lender or insurance underwriter → financial. When genuinely mixed, follow how the **precedent
deals themselves were valued** (what multiple the PRs quote) and how equity research covers the
sector. Do not mix families within one table.

## Source ladder — per figure, top-down

Work the ladder top-down for each deal's figures; stop at the first rung that produces a usable
value.

### Rung 1 — Disclosed multiple (preferred)

If the deal source quotes the multiple directly — "~12.5x LTM EBITDA", "2.8x revenue", "1.6x
book value", "~18x P/E" — write that number as a **literal into the multiple column** (S–Z),
overwriting the template's ratio formula. A directly disclosed multiple is the cleanest figure:
it's exactly what the buyer/seller quoted and carries no TEV-measurement, FX, or period-mismatch
noise. The group / global statistic rows average the literal in automatically.

Look in: the acquiror's (or target's) deal press release; an 8-K / 6-K / 40-F exhibit or
investor deck filed with the announcement; the deal-day conference-call transcript; or major
financial news quoting deal-day materials.

"~Nx" / "approximately Nx" still counts as a disclosed multiple — use it. Wanting to verify
against filings is fine, but the cell value is the disclosed multiple.

### Rung 2 — Disclosed $ metric (most-recent proxy)

If no multiple is disclosed but the absolute figure is (LTM revenue of $X, FY net income of
$Y, book value of $Z), write the **source-currency $MM** value into the input column (K–R). The
template's ratio formula converts it (`× C` FX) and divides by TEV on refresh.

When a clean LTM (or NTM) figure isn't disclosed, **use the most recent reported figure as the
LTM/NTM proxy** — the latest full-year or latest reported LTM-ish number from the deal sources
or the target's most recent filing. Do **not** stitch LTM together from three filings (MRQ +
prior FY − prior-year stub); that route is intentionally not used here. Note any proxy in the
summary.

### Rung 3 — Blank

If neither a multiple nor a usable $ figure is disclosed from a reputable source, leave the
cell blank. Never fabricate or estimate a figure. A deal with a disclosed TEV but a missing
metric is still useful — keep it; just leave that metric blank.

## "Disclosed" — precise definition

A figure or multiple is "disclosed" only if it appears in: the target's or acquiror's deal
press release; an 8-K / 6-K / 40-F exhibit or investor deck filed with the announcement; the
deal-day conference-call transcript; or major news (Bloomberg, Reuters, WSJ, FT, Globe and
Mail, Financial Post, S&P Global) **quoting deal-day materials**. Numbers computed post-hoc by
analysis blogs / deal-recap sites are *implied by a third party*, not disclosed — do not use
them.

## "Pro forma" almost never means synergies

In deal-source multiples, "pro forma" almost always means pro forma for divestitures or
continuing operations — not pro forma for buyer synergies. Synergies are virtually always
called out separately ("$X of expected run-rate cost savings"), not embedded in the headline
multiple. Treat a multiple as synergy-inclusive only if the source explicitly says "including
synergies", "post-synergy", or "synergized".

## Currency

Enter each deal's figures in the **currency of the original source**, with the ISO-3 code in
column B. The column-C CapIQ FX formula converts source → output currency (`C2`), and the ratio
formulas apply that conversion (`(metric × C) / TEV`). A **disclosed multiple** written into
S–Z is dimensionless — it needs no FX and inherits nothing; just write the number. TEV (column
I) is also entered in source currency; column J (`=+I*C`) converts it.

## Transaction-selection criteria

- **Comparable target** — sector, business model, client segment, asset class, scale similar to
  the input company. Comparability comes first.
- **Disclosed TEV — required.** Never include a deal with no deal value; the row is unusable
  for multiple analysis without it.
- **Recency** — prefer deals announced/closed in the last ~6–8 years.
- **At least one metric — required.** Each kept deal **must** yield at least one of the family's
  multiples (from a disclosed multiple or a disclosed $ figure); a TEV-only row with no metric is
  rejected by the builder. Drop and replace a deal you can't value rather than shipping an empty
  row.
- Organise into two peer groups and **aim to fill all six rows in each** (12 deals total, e.g. by
  sub-sector or deal type). A single group is acceptable only when a credible second peer group
  genuinely doesn't exist; otherwise fill both. Leave an unused group as its `[Group #N]`
  placeholder.

## Reputable sources — hyperlink targets

Hyperlink each figure (AB–AG) to a **primary** source. Prefer:

- the target's or acquiror's investor-relations / press-release page;
- `sec.gov` (EDGAR); `sedarplus.ca` / `sedar.com` (Canadian filings);
- major financial news: `bloomberg.com`, `reuters.com`, `wsj.com`, `ft.com`,
  `theglobeandmail.com`, `financialpost.com`, `spglobal.com`;
- the issuer-PR wires `businesswire.com`, `globenewswire.com`, `prnewswire.com` — only when
  carrying the verbatim issuer press release.

Avoid analyst blogs and deal-recap sites (e.g. mergersight, eResearch, marketscreener,
tipranks, growjo, Substack/Medium, LinkedIn posts) as the primary source for any figure — their
numbers are third-party-implied, not disclosed. If the only source for a figure is off-list,
keep searching primary sources or leave the cell blank; don't cite the off-list source.
