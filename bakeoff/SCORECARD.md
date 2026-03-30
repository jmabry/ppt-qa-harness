# Bakeoff Scorecard

## Conditions

All 9 decks were generated in a single pass using `claude --print` with each skill's docs injected as system prompt. **No QA loops ran** — subagents were denied Bash permission, so no skill got to render, inspect, or fix its output. Scores reflect first-pass generation quality only.

Decks were rendered to JPEG via LibreOffice at 120 DPI and visually inspected by a fresh subagent per prompt group (no knowledge of which skill produced which deck). Text extraction via `markitdown` supplemented visual inspection for content accuracy checks.

**Scoring bias disclosure:** The deck-builder skill is this repo's own skill. Visual inspection was delegated to fresh subagents blind to skill identity. These results are honest: deck-builder performed worst in this run due to systematic layout failures that a QA loop would catch.

## Scoring Methodology

Each deck is rated 1-5 on 5 dimensions (3 universal + 2 prompt-specific). Scores are based on visual inspection of rendered slide images (LibreOffice PDF export, 120 DPI JPEG).

| Score | Meaning                                              |
| ----- | ---------------------------------------------------- |
| **5** | Excellent — polished, no issues, would present as-is |
| **4** | Good — minor issues that don't hurt usability        |
| **3** | Adequate — noticeable issues but still functional    |
| **2** | Below average — significant issues hurt usability    |
| **1** | Poor — broken rendering, missing content, unusable   |

### Universal Dimensions

| Dimension              | 1 (Poor)                                      | 3 (Adequate)                           | 5 (Excellent)                                          |
| ---------------------- | --------------------------------------------- | -------------------------------------- | ------------------------------------------------------ |
| **Visual appeal**      | Plain text on white, no design effort         | Consistent colors, some layout variety | Cohesive palette, varied layouts, professional polish  |
| **Layout correctness** | Text overflow, clipping, overlapping elements | Minor alignment issues, mostly clean   | No overflow, consistent margins, proper whitespace     |
| **Readability**        | Tiny fonts, low contrast, cramped spacing     | Readable but some dense sections       | Clear hierarchy, good contrast, comfortable font sizes |

### Prompt-Specific Dimensions

| Prompt                              | Dimension 1                                                                   | Dimension 2                                                       |
| ----------------------------------- | ----------------------------------------------------------------------------- | ----------------------------------------------------------------- |
| **corporate** (United Airlines)  | **Content accuracy** — real financials, plausible data, internal consistency   | **Data density** — charts, tables, KPI cards, metrics per slide  |
| **software** (Microservices)     | **Technical accuracy** — correct terminology, sound architecture              | **Data visualization** — diagrams, charts, flow representation  |
| **strategy** (Board Review)      | **Executive presence** — would you show this to a board?                      | **Data density** — charts, tables, KPI cards, metrics             |

---

## Corporate (United Airlines Investor Update)

This is the most data-dense prompt — a 16-18 slide institutional investor deck with specific financials, fleet data, and strategic metrics.

| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                            |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 4            | 3         | 4         | DB: clean navy/gold, professional. MiniMax: systematic right-column clipping undermines the dark theme. Anthropic: consistent, typographically strong across 18 slides. |
| Content accuracy   | 2            | 3         | 4         | DB: fuel sensitivity callout states "$0.10/gal ≈ $40M" — arithmetic wrong by 10x (should be ~$400M at 4B+ gallon burn). MiniMax: minor rounding on FY2022; fleet counts use "operated" vs "total" methodology, unlabeled. Anthropic: internally consistent throughout; all key figures cross-check. |
| Layout correctness | 3            | 2         | 4         | DB: title text overlap on slide 1, appendix text overwrites hub table on slide 16, persistent 40-50% blank lower halves on 8+ slides. MiniMax: systematic right-column clipping on slides 4, 5, 6, 7, 9, 14; title slide renders 3 text lines stacked on top of each other. Anthropic: mild overflow bottom of slide 2, minor footnote clipping on slides 11/18. |
| Data density       | 2            | 4         | 5         | DB: consistently 40-50% blank space; text-only product slide (slide 8); thin single-table slides (5, 6, 11). MiniMax: dense throughout with bar charts, progress bars, additional metrics panels. Anthropic: 18 slides of high-density tables, KPI cards, analytical text — highest information volume. |
| Readability        | 4            | 3         | 4         | DB: clear hierarchy, comfortable fonts, easy to scan. MiniMax: slide 10 text too small; clipping destroys readability in affected columns. Anthropic: hierarchy clear; footnotes occasionally too small (slides 5, 13, 18). |
| **Total**          | **15/25**    | **15/25** | **21/25** |                                                                                                                                                                  |

## Software (Monolith to Microservices)

6-slide technical presentation for engineering leadership. Every slide should have tables, diagrams, or metrics.

| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                                    |
| ------------------ | ------------ | --------- | --------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Visual appeal      | 3            | 3         | 5         | DB: functional dark blue/white, flat. MiniMax: good stat boxes and architecture diagram but rough slides 4-5. Anthropic: cohesive dark/cyan palette, varied layouts (hero, ADR columns, color-coded tables). |
| Technical accuracy | 4            | 5         | 5         | All three correctly apply strangler fig, DORA metrics, service mesh, Kafka, saga. MiniMax adds merge conflict + rollback metrics and quantified Phase 3 confidence (65%/40%) — richer context. |
| Layout correctness | 2            | 2         | 4         | DB: slide 5 catastrophically broken — section header renders as "OBERSMEBILITY STACK" (font corruption), "CRITICAL" badge splits as "CRITIC AL", right-side monitoring table clipped. MiniMax: slides 4-5 have severe text overlap and overprinting; slide 6 footer overwrites squad table. Anthropic: slide 5 has one minor wrap ("Observabilit Y"), slide 1 footer/watermark collision at bottom edge. |
| Data visualization | 3            | 4         | 4         | MiniMax slide 2 has the best architecture diagram — hierarchical layout showing gateway → services → Kafka → data stores. DB has a simple linear box chain. Anthropic has no diagram but uses well-structured color-coded tables with severity badges. |
| Readability        | 3            | 2         | 5         | Anthropic: strong contrast, comfortable fonts throughout. DB: mostly readable except slide 5 is broken. MiniMax: slides 4-5 dense and garbled; slide 6 footer overprints squad table. |
| **Total**          | **15/25**    | **16/25** | **23/25** |                                                                                                                                                                          |

## Strategy (Q3 Board Review)

6-slide board-level presentation for Series B SaaS company. Board members read ahead — optimize for information density.

| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 3            | 4         | 5         | DB: clean theme but catastrophic rendering on slides 2-5 undermines it. MiniMax: polished dark-navy/teal, well-organized. Anthropic: cohesive palette with left-border accent motif, varied slide treatments. |
| Executive presence | 3            | 4         | 5         | DB: slide 1 is strong (3 wins/2 concerns/1 decision), but slides 2-5 are broken — board would not receive this deck. MiniMax: full board asks with CEO attribution, confidential footer, 3-column ask layout. Anthropic: CONFIDENTIAL on every slide, ROI on each ask, at-risk account tables, Series C option analysis. |
| Layout correctness | 1            | 4         | 3         | DB: 4 of 6 slides (2, 3, 4, 5) have severe vertical text overflow — table column headers render character-by-character in a vertical stack, making entire table payload unreadable. Also missing third board ask on slide 6. MiniMax: clean with no overflow detected. Anthropic: slides 4-5 have right-edge column truncation ("Preventabl e?" header wrap, Impact column clipped). |
| Data density       | 2            | 4         | 5         | DB: content is present in source but unreadable due to rendering failures. MiniMax: complete dataset — competitive win rates, at-risk accounts, ARR trend. Anthropic: all KPIs, P&L, pipeline, churn breakdown, competitive table, Series C options — most complete. |
| Readability        | 1            | 4         | 4         | DB: slides 2-5 are effectively unreadable. MiniMax: slide 4 is quite small/zoomed-out but otherwise good contrast and font sizes. Anthropic: minor truncation on 2 slides but majority is excellent. |
| **Total**          | **10/25**    | **20/25** | **22/25** |                                                                                                                                                      |

---

## Overall

| Skill            | corporate | software | strategy | Grand Total |
| ---------------- | --------- | -------- | -------- | ----------- |
| **Anthropic**    | 21/25     | 23/25    | 22/25    | **66/75**   |
| **MiniMax**      | 15/25     | 16/25    | 20/25    | **51/75**   |
| **deck-builder** | 15/25     | 15/25    | 10/25    | **40/75**   |

**Anthropic wins this run by a large margin (66 vs 51 vs 40). deck-builder came last due to systematic layout failures across all three prompts that a QA loop would have caught.**

---

## Skill Comparison

### Anthropic — best first-pass output, highest data density

**Strengths:** Consistently strong across all three prompts. Best visual polish — dark/light sandwich structure, left-border accent motif, colored severity badges. Highest data density: 18-slide UAL deck with two appendix slides, all packed. Strongest writing — analytical narrative framing, specific context boxes on every slide. Proper financial formatting (parentheses for negatives, YoY columns). Content accuracy highest — numbers cross-check throughout.

**Weaknesses:** Minor right-edge column truncation on 2 slides in strategy (roadmap Impact column, competitive table Zenith column). Some footnotes at 8pt are difficult to read. Footer occasionally merges unrelated topics on one line (slide 11).

### MiniMax — most visually ambitious, inconsistent layout

**Strengths:** Best architecture diagram in software deck (hierarchical, not just a linear box chain). Good data density — additional metrics (merge conflicts, rollback rate, test flakiness) not present in other decks. Page badges on every non-title slide. Competitive win rates table on strategy slide 5. Cover slides use full KPI dashboard layout. Dark themes are well-executed when layout holds.

**Weaknesses:** Systematic right-column clipping on 6+ slides in the corporate deck — FY2026E column cut off on slides 4, 6, 7, 9, 14. Title slide text stack collision (corporate slide 1: three text lines rendered on top of each other). Software slides 4-5 have severe overlap and overprinting.

### deck-builder — worst first-pass quality, QA-dependent

**Strengths:** Clean, consistent palette and typography when rendering is correct. Readable font sizes. Slide 1 exec summaries are well-structured. Financial data is accurately transcribed (except fuel sensitivity error). Charts (where present) are clean.

**Weaknesses:** **Severe layout failures across all three prompts** — these would all be caught and fixed by the mandatory QA loop, but without it the output is not presentable:
- Strategy slides 2-5: table column headers render as vertical character stacks (each column letter on its own line) — entire table content unreadable on 4 of 6 slides
- Software slide 5: font corruption producing "OBERSMEBILITY STACK" and mid-word breaks ("CRITIC AL", "Impac t") — same charSpacing bug noted in prior runs
- Corporate: fuel sensitivity arithmetic error (states $40M, should be ~$400M), title text overlap, persistent 40-50% blank space on 8+ slides, appendix text overwrites table

---

## Key Issues Found

### Critical (rendering bugs)

| Skill        | Deck          | Slide(s) | Issue                                                                 |
| ------------ | ------------- | -------- | --------------------------------------------------------------------- |
| deck-builder | Strategy      | 2, 3, 4, 5 | Table column headers render as vertical character stacks — entire table payload unreadable on 4 slides |
| deck-builder | Strategy      | 6        | Third board ask (Series C) missing from slide                        |
| deck-builder | Software      | 5        | Font corruption: "OBERSMEBILITY STACK", "CRITIC AL", "Impac t", right-side table clipped |
| deck-builder | Corporate     | 5        | Fuel sensitivity states "$40M" — 10x arithmetic error (should be ~$400M at 4B+ gallon burn) |
| deck-builder | Corporate     | 1        | Title text overlaps subtitle heading                                 |
| MiniMax      | Corporate     | 1        | Title slide: 3 text lines rendered stacked on top of each other       |
| MiniMax      | Corporate     | 4, 6, 7, 9, 14 | Systematic right-column clipping — FY2026E data lost on 5 slides   |
| MiniMax      | Software      | 4        | Phase 2 and Phase 3 columns have severe text overlap and overprinting |
| MiniMax      | Software      | 5        | Mitigation column garbled; observability table clipped               |
| Anthropic    | Strategy      | 4, 5     | Rightmost columns truncated (Roadmap Impact, Competitive table)      |

### Moderate

| Skill        | Deck          | Slide(s) | Issue                                                          |
| ------------ | ------------- | -------- | -------------------------------------------------------------- |
| deck-builder | Corporate     | multiple | 40-50% blank lower halves on 8+ slides (density problem)       |
| deck-builder | Corporate     | 8        | Product slide is text-only, no data or metrics                 |
| deck-builder | Corporate     | 16       | Appendix workforce text overwrites hub table rows              |
| MiniMax      | Corporate     | 3        | Fleet counts labeled "operated" vs. "total" — methodology mismatch, unlabeled |
| MiniMax      | Corporate     | 10       | New products quadrant text too small, bleeds across card borders |
| MiniMax      | Software      | 6        | "Principle" footer text overprints SRE row in team table       |
| Anthropic    | Corporate     | 2        | Investment thesis body text clipped at slide bottom (last sentence) |
| Anthropic    | Software      | 5        | "Observabilit Y" word-wrap in narrow Phase 2 type column       |
| Anthropic    | Strategy      | 3        | "Preventabl e?" header wrap in churn table                     |

---

## Improvements to make (deck-builder)

The core problem is that first-pass generation has severe layout failures that only the QA loop catches. These are not cosmetic — they make slides non-presentable. Specific fixes:

- **Fix the vertical column stack bug** — when table `colW` array values don't sum to `w`, columns collapse to near-zero width and text renders vertically. Add a validation helper that asserts `sum(colW) ≈ w` before generating any table.
- **Fix the charSpacing font corruption** — charSpacing on any text object (including section labels) causes garbled rendering in LibreOffice. Remove charSpacing from all section labels; use letter-spacing only on shapes with no text dependency.
- **Fill vertical space** — biggest visual gap vs. competitors. 40-50% blank lower halves on multiple slides. Add a `checkVerticalFill(slide, y)` helper that warns when less than 70% of CONTENT_H is used.
- **Chart axis labels** — always include units ($B, %, ¢) on axes and data labels.
- **Arithmetic validation** — add commented assertions next to any calculated figures ($40M should be $400M).
- **Table colW validation** — this is the root cause of the strategy and software table failures. Critical fix.

## Improvements to make (all skills)

- **MiniMax corporate right-column clipping** — the FY2026E column disappears on 5 slides. Root cause likely the column width not accounting for the right margin when the deck is narrow. Use a colW array that sums to exactly `w - 2*x`.
- **All skills** — footnotes and fine-print at the slide bottom are consistently near-illegible when projected. 8pt minimum for footnotes is insufficient — 9pt minimum.
- **All skills** — competitive landscape tables get truncated when there are 4+ competitors. Consider dropping to 3 columns or using two rows per competitor.
