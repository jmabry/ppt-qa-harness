# Bakeoff Scorecard

## Conditions

All 9 decks were generated in a single pass using `claude --print` with each skill's docs injected as system prompt. **No QA loops ran** — subagents were denied Bash permission, so no skill got to render, inspect, or fix its output. Scores reflect first-pass generation quality only.

## Scoring Methodology

Each deck is rated 1-5 on 5 dimensions (3 universal + 2 prompt-specific). Scores are based on visual inspection of rendered slide images (LibreOffice PDF export, 150 DPI JPEG).

| Score | Meaning |
|-------|---------|
| **5** | Excellent — polished, no issues, would present as-is |
| **4** | Good — minor issues that don't hurt usability |
| **3** | Adequate — noticeable issues but still functional |
| **2** | Below average — significant issues hurt usability |
| **1** | Poor — broken rendering, missing content, unusable |

### Universal Dimensions

| Dimension | 1 (Poor) | 3 (Adequate) | 5 (Excellent) |
|-----------|----------|--------------|---------------|
| **Visual appeal** | Plain text on white, no design effort | Consistent colors, some layout variety | Cohesive palette, varied layouts, professional polish |
| **Layout correctness** | Text overflow, clipping, overlapping elements | Minor alignment issues, mostly clean | No overflow, consistent margins, proper whitespace |
| **Readability** | Tiny fonts, low contrast, cramped spacing | Readable but some dense sections | Clear hierarchy, good contrast, comfortable font sizes |

### Prompt-Specific Dimensions

| Prompt | Dimension 1 | Dimension 2 |
|--------|-------------|-------------|
| **01-creative** (Pasta) | **Content quality** — accuracy, completeness, engaging tone | **Slide variety** — different layouts across slides vs repetitive |
| **02-software** (Microservices) | **Technical accuracy** — correct terminology, sound architecture | **Data visualization** — diagrams, charts, flow representation |
| **03-strategy** (Board Review) | **Executive presence** — would you show this to a board? | **Data density** — charts, tables, KPI cards, metrics |

---

## Prompt 1: Creative (Homemade Pasta)

| Dimension | deck-builder | MiniMax | Anthropic | Notes |
|---|---|---|---|---|
| Visual appeal | 4 | 4 | 4 | All three chose warm earth tones. DB and MiniMax use decorative shapes; Anthropic uses split-panel layouts. |
| Content quality | 5 | 4 | 5 | DB and Anthropic have specific, engaging copy with personality. MiniMax is accurate but slightly more generic. |
| Layout correctness | 4 | 3 | 4 | MiniMax slide 4 has text overflow past card boundary; slides 4+6 have low-contrast body text. DB has minor right-side cramping on shape slides. |
| Slide variety | 4 | 4 | 4 | All use 4-5 distinct layouts. DB slides 4-6 share a template (intentional — one per pasta shape). MiniMax varies more between shape slides. |
| Readability | 4 | 3 | 4 | MiniMax loses points for low-contrast tan-on-white body text on slides 4 and 6. DB and Anthropic are consistently readable. |
| **Total** | **21/25** | **18/25** | **21/25** | |

## Prompt 2: Software (Monolith to Microservices)

| Dimension | deck-builder | MiniMax | Anthropic | Notes |
|---|---|---|---|---|
| Visual appeal | 4 | 3 | 4 | DB dark navy+teal is polished. Anthropic dark theme with colored accents is professional. MiniMax navy+gold is functional but monotone. |
| Technical accuracy | 5 | 5 | 5 | All three correctly apply strangler fig, DORA metrics, service mesh, event-driven architecture. |
| Layout correctness | 3 | 3 | 3 | DB has garbled text on slide 7 (corrupted heading). MiniMax has header text wrapping ("LIKELIHOO/D") and Gantt chart clipping. Anthropic has text truncation on slide 6 ("Foundatio n"). |
| Data visualization | 4 | 3 | 3 | DB has a clean architecture diagram, phased timeline bar, and org chart. MiniMax has a Gantt chart (ambitious but clipped). Anthropic architecture diagram is basic; no actual charts. |
| Readability | 3 | 4 | 3 | DB dark theme has small text + garbled heading. MiniMax yellow-on-navy is high contrast. Anthropic dark theme has small text and truncation. |
| **Total** | **19/25** | **18/25** | **18/25** | |

## Prompt 3: Strategy (Q3 Board Review)

| Dimension | deck-builder | MiniMax | Anthropic | Notes |
|---|---|---|---|---|
| Visual appeal | 5 | 4 | 4 | DB is the most polished — executive-grade navy theme, consistent KPI cards, clean charts. MiniMax and Anthropic are professional but DB edges ahead. |
| Executive presence | 5 | 5 | 5 | All three produce board-ready decks with exec summary, specific asks, confidential markings, and financial appendices. |
| Layout correctness | 5 | 4 | 4 | DB has zero rendering bugs. MiniMax slide 12 has a truncated cash flow table. Anthropic slides 5+9 have charts without legends. |
| Data density | 5 | 5 | 5 | All three pack KPI cards, bar/line charts, tables, pipeline breakdowns, and full P&L appendices. Excellent across the board. |
| Readability | 4 | 4 | 4 | All three are dense but appropriate for board materials. Small text in some tables/chart legends across all skills. |
| **Total** | **24/25** | **22/25** | **22/25** | |

---

## Overall

| Skill | 01-creative | 02-software | 03-strategy | Grand Total |
|---|---|---|---|---|
| **deck-builder** | 21/25 | 19/25 | 24/25 | **64/75** |
| **MiniMax** | 18/25 | 18/25 | 22/25 | **58/75** |
| **Anthropic** | 21/25 | 18/25 | 22/25 | **61/75** |

---

## Key Issues Found

### Critical (rendering bugs)

| Skill | Deck | Slide | Issue |
|-------|------|-------|-------|
| deck-builder | Microservices | 7 | Garbled/corrupted heading text ("A_N_E_R_T_H_R_G_E_S...") |
| MiniMax | Pasta | 4 | Text overflows card boundary; body text near-invisible (tan on white) |
| MiniMax | Microservices | 6 | Gantt chart "Monolith Retire" clipped at right edge |
| Anthropic | Microservices | 6 | Timeline headings truncated mid-word ("Foundatio n", "Accelerat e") |

### Moderate

| Skill | Deck | Slide | Issue |
|-------|------|-------|-------|
| MiniMax | Pasta | 6 | Low-contrast body text in all three step cards |
| MiniMax | Microservices | 5 | "LIKELIHOOD" header wraps to "LIKELIHOO/D" |
| MiniMax | Board Review | 12 | Cash Flow table truncated at bottom |
| Anthropic | Board Review | 5, 9 | Charts with multiple series but no legend |

---

## Observations

### What deck-builder did well
- Best board review deck — zero rendering bugs, highest visual polish, consistent KPI + chart + table pattern
- Strong visual variety across all three decks (cards, tables, charts, diagrams, org charts)
- Architecture diagram and phased timeline bar in microservices deck are effective custom visualizations
- Warm, intentional color palettes matched to each topic

### What MiniMax did well
- Excellent board-level content — the exec summary "3 wins / 2 concerns / 1 decision" format is textbook
- Best technical content in microservices (DORA metrics, O(n^2) coordination note)
- Consistent page number badges on every non-title slide
- Ambitious Gantt chart attempt in microservices (most complex visualization of any deck)

### What Anthropic did well
- Strongest writing — pasta deck copy has real personality and specific technique advice
- Consistent design systems within each deck (matching title/closing slides, coherent accents)
- Board review has proper financial formatting (parentheses for negatives, YoY column)
- Two-panel split layouts are distinctive and effective

### Improvements to make (deck-builder)
- Fix charSpacing bug that caused garbled heading on microservices slide 7
- Add chart axis labels and legends consistently
- Consider page number badges (MiniMax does this well)
- Slightly increase body text size on dark-background slides
