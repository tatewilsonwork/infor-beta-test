# Historical-regression fixtures

One deck per historical bug, so `deck_contract.verify_deck` keeps proving it
catches them. Wired up in `scripts/tests/test_deck_contract.py`.

**All three are real replays, not synthesized.** Each is the analyst's actual run
output with every slide except the defective one deleted (`pptx_helpers.delete_slide`
— which also drops the slide part's relationship, so unreferenced media is not
serialised). The surviving slide's shape XML, the slide layouts, the masters, and
the embedded fonts are byte-for-byte the original. Nothing was authored, moved, or
retyped to provoke a finding.

The strip was verified not to change the measurement:

| Fixture | Finding | Source deck | Stripped deck |
|---|---|---|---|
| `prl17-market-entry-table.pptx` | `rendered-overflow` | 0.180" | 0.180" |
| `prl18-risk-table.pptx` | `table-taller-than-library` | 5.360" | 5.360" |
| `prl14-overview-bullets.pptx` | `rendered-overflow` | 2.580" | 2.580" |

## Provenance

| Fixture | Source run | Kept slide | Bug | Fixed in |
|---|---|---|---|---|
| `prl17-market-entry-table.pptx` | `Project PRL17`, 2026-07-09 | 14 — `Potential Mexican Market Entry Targets (1 of 4)` | Table declares 5.710" and **renders 5.91"** | v0.5.23 |
| `prl18-risk-table.pptx` | `Project PRL18`, 2026-07-10 | 10 — `Potential Perceived Acquiror Considerations and Mitigants` | Table **declares** 5.360" vs the library's 5.1715" | v0.5.24 |
| `prl14-overview-bullets.pptx` | `Project PRL14`, 2026-06-30 | 7 — `Introduction to Propel Holdings, Inc.` | Overview bullets render through the LTM Revenue Breakdown band | v0.5.23 |

### Why PRL14 rather than PRL17 for the overview overflow

The v0.5.23 CHANGELOG entry attributes the overview-overflow fix to the PRL17
run, but the PRL17 folder on disk holds a **post-fix re-run**: its overview
`TextBox 9` carries an explicit `normAutofit fontScale="85000"` and a box sized to
the band (bottom 4.64", clear of the 4.747" band header). The pre-fix PRL17
artefact was overwritten.

PRL14 is a genuine pre-fix artefact of the identical defect — same deck, same
slide, same code path, one release earlier. Its signature is the pre-v0.5.23 one:

- `TextBox 9` declared 1.47"–2.05" (0.58" tall) holding 7 bullets / 1,126 chars
- **no `normAutofit` element at all**, so nothing shrinks on open
- the LTM Revenue Breakdown band header (`Rectangle 15`) at 4.747"

Every pre-v0.5.23 pitch run on disk (PRL2, 4, 5, 7, 10–15) shares it. PRL14 is the
most recent, and the one whose copy is long enough to overflow visibly: its last
bullet renders on top of the navy band.

## Why these carry embedded fonts (~900 KB of each file)

Most of each fixture's ~1.2 MB is `ppt/fonts/*.fntdata` — the deck embeds
Palatino Linotype and friends. They are deliberately **not** stripped: the
`rendered-overflow` measurement is a text-layout measurement, and removing the
embedded fonts would make it depend on which fonts the host has installed. That is
precisely the fragility Phase A identified (LibreOffice has no
`palatinolinotype` substitution entry, so a Linux box without `fonts-urw-base35`
falls back to Times/DejaVu metrics). Keeping them keeps the fixtures measuring the
same thing everywhere.

## Do not "fix" these decks

They are defect specimens. A finding disappearing here means the contract
regressed, not that the fixture needs updating.

> **These are INFOR client work product, and this repository is public.** Every
> slide is marked *Private & Confidential* by the master, and the market-entry
> fixture names specific Mexican acquisition targets. They are committed here on
> the precedent set in Phase A, which put the full `pitch-deck.pptx` — containing
> the same engagement, the same target names, and the same slide concepts — in
> `fixtures/`, so these add no disclosure beyond what is already published. That
> is a reason not to treat the decision as reopened, **not** a reason to add more:
> weigh any further client artefact on its own merits before committing it, and
> raise it with the analyst if it introduces anything new.
