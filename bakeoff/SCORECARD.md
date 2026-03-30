# Bakeoff Scorecard

## Conditions

All 12 decks were generated in a single pass using `claude --print` with each skill's docs injected as system prompt. **No QA loops ran** — subagents were denied Bash permission, so no skill got to render, inspect, or fix its output. Scores reflect first-pass generation quality only.

## Scoring Methodology

Each deck is rated 1-5 on 5 dimensions (3 universal + 2 prompt-specific). Scores are based on visual inspection of rendered slide images (LibreOffice PDF export, 150 DPI JPEG).


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


| Prompt                              | Dimension 1                                                      | Dimension 2                                                       |
| ----------------------------------- | ---------------------------------------------------------------- | ----------------------------------------------------------------- |
| **00-corporate** (United Airlines)  | **Content accuracy** — real financials, plausible data, internal consistency | **Data density** — charts, tables, KPI cards, metrics per slide |
| **01-creative** (Pasta)             | **Content quality** — accuracy, completeness, engaging tone      | **Slide variety** — different layouts across slides vs repetitive |
| **02-software** (Microservices)     | **Technical accuracy** — correct terminology, sound architecture | **Data visualization** — diagrams, charts, flow representation   |
| **03-strategy** (Board Review)      | **Executive presence** — would you show this to a board?         | **Data density** — charts, tables, KPI cards, metrics             |


---

## Prompt 0: Corporate (United Airlines Investor Update) ★

This is the most representative real-world test — a data-dense corporate investor deck generated from a long-form research document with specific financials, fleet data, and strategic metrics.


| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                            |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 3            | 4         | 4         | DB has consistent navy/gold but excessive whitespace on many slides. MiniMax has clean card-based design with decorative shapes. Anthropic has bold dark/light sandwich. |
| Content accuracy   | 5            | 5         | 4         | DB and MiniMax data is internally consistent and accurate. Anthropic has a MAX 8 arithmetic error (215 firm orders but 142+113=255 remaining).                     |
| Layout correctness | 4            | 3         | 3         | DB has no overflow/clipping but wastes vertical space. MiniMax has text wrapping bugs (EWR split, % signs). Anthropic has pervasively small text across 12+ slides. |
| Data density       | 3            | 4         | 4         | DB is chart-light (6/16 slides have charts) with repetitive KPI+bullets pattern. MiniMax has KPI cards, tables, charts, progress bars on most slides. Anthropic packs charts + tables on every data slide. |
| Readability        | 4            | 3         | 3         | DB text is readable but some chart labels lack units. MiniMax slide 15 is nearly illegible, slides 5/13 have low-contrast footnotes. Anthropic has pervasively small sub-text and footnotes. |
| **Total**          | **19/25**    | **19/25** | **18/25** |                                                                                                                                                                  |


## Prompt 1: Creative (Homemade Pasta)


| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                           |
| ------------------ | ------------ | --------- | --------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 4            | 4         | 4         | All three chose warm earth tones. DB and MiniMax use decorative shapes; Anthropic uses split-panel layouts.                                     |
| Content quality    | 5            | 4         | 5         | DB and Anthropic have specific, engaging copy with personality. MiniMax is accurate but slightly more generic.                                  |
| Layout correctness | 4            | 3         | 4         | MiniMax slide 4 has text overflow past card boundary; slides 4+6 have low-contrast body text. DB has minor right-side cramping on shape slides. |
| Slide variety      | 4            | 4         | 4         | All use 4-5 distinct layouts. DB slides 4-6 share a template (intentional — one per pasta shape). MiniMax varies more between shape slides.     |
| Readability        | 4            | 3         | 4         | MiniMax loses points for low-contrast tan-on-white body text on slides 4 and 6. DB and Anthropic are consistently readable.                     |
| **Total**          | **21/25**    | **18/25** | **21/25** |                                                                                                                                                 |


## Prompt 2: Software (Monolith to Microservices)


| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                                                    |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 4            | 3         | 4         | DB dark navy+teal is polished. Anthropic dark theme with colored accents is professional. MiniMax navy+gold is functional but monotone.                                                  |
| Technical accuracy | 5            | 5         | 5         | All three correctly apply strangler fig, DORA metrics, service mesh, event-driven architecture.                                                                                          |
| Layout correctness | 3            | 3         | 3         | DB has garbled text on slide 7 (corrupted heading). MiniMax has header text wrapping ("LIKELIHOO/D") and Gantt chart clipping. Anthropic has text truncation on slide 6 ("Foundatio n"). |
| Data visualization | 4            | 3         | 3         | DB has a clean architecture diagram, phased timeline bar, and org chart. MiniMax has a Gantt chart (ambitious but clipped). Anthropic architecture diagram is basic; no actual charts.   |
| Readability        | 3            | 4         | 3         | DB dark theme has small text + garbled heading. MiniMax yellow-on-navy is high contrast. Anthropic dark theme has small text and truncation.                                             |
| **Total**          | **19/25**    | **18/25** | **18/25** |                                                                                                                                                                                          |


## Prompt 3: Strategy (Q3 Board Review)


| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 5            | 4         | 4         | DB is the most polished — executive-grade navy theme, consistent KPI cards, clean charts. MiniMax and Anthropic are professional but DB edges ahead. |
| Executive presence | 5            | 5         | 5         | All three produce board-ready decks with exec summary, specific asks, confidential markings, and financial appendices.                               |
| Layout correctness | 5            | 4         | 4         | DB has zero rendering bugs. MiniMax slide 12 has a truncated cash flow table. Anthropic slides 5+9 have charts without legends.                      |
| Data density       | 5            | 5         | 5         | All three pack KPI cards, bar/line charts, tables, pipeline breakdowns, and full P&L appendices. Excellent across the board.                         |
| Readability        | 4            | 4         | 4         | All three are dense but appropriate for board materials. Small text in some tables/chart legends across all skills.                                  |
| **Total**          | **24/25**    | **22/25** | **22/25** |                                                                                                                                                      |


---

## Overall


| Skill            | 00-corporate | 01-creative | 02-software | 03-strategy | Grand Total  |
| ---------------- | ------------ | ----------- | ----------- | ----------- | ------------ |
| **deck-builder** | 19/25        | 21/25       | 19/25       | 24/25       | **83/100**   |
| **MiniMax**      | 19/25        | 18/25       | 18/25       | 22/25       | **77/100**   |
| **Anthropic**    | 18/25        | 21/25       | 18/25       | 22/25       | **79/100**   |


---

## Key Issues Found

### Critical (rendering bugs)


| Skill        | Deck          | Slide | Issue                                                                 |
| ------------ | ------------- | ----- | --------------------------------------------------------------------- |
| deck-builder | Microservices | 7     | Garbled/corrupted heading text ("A_N_E_R_T_H_R_G_E_S...")             |
| MiniMax      | Pasta         | 4     | Text overflows card boundary; body text near-invisible (tan on white) |
| MiniMax      | Microservices | 6     | Gantt chart "Monolith Retire" clipped at right edge                   |
| MiniMax      | UAL           | 8, 10 | EWR airport code split across lines; % signs wrap to next line        |
| MiniMax      | UAL           | 15    | ESG slide text nearly illegible (faded/small bullets in all 3 cards)  |
| Anthropic    | Microservices | 6     | Timeline headings truncated mid-word ("Foundatio n", "Accelerat e")   |
| Anthropic    | UAL           | 7     | MAX 8 arithmetic error (215 orders but 142+113=255 delivered+remaining) |


### Moderate


| Skill        | Deck          | Slide    | Issue                                                  |
| ------------ | ------------- | -------- | ------------------------------------------------------ |
| deck-builder | UAL           | multiple | Excessive whitespace — content fills only 60-70% of slide on 5+ slides |
| deck-builder | UAL           | 5, 7     | Chart axis labels missing units ($B, %)                |
| MiniMax      | Pasta         | 6        | Low-contrast body text in all three step cards         |
| MiniMax      | Microservices | 5        | "LIKELIHOOD" header wraps to "LIKELIHOO/D"             |
| MiniMax      | Board Review  | 12       | Cash Flow table truncated at bottom                    |
| MiniMax      | UAL           | 12       | Leverage chart vs KPI card data inconsistency (3x vs 2.6x) |
| Anthropic    | Board Review  | 5, 9     | Charts with multiple series but no legend              |
| Anthropic    | UAL           | multiple | Pervasively small sub-text and footnotes (12+ slides)  |
| Anthropic    | UAL           | 6, 9, 13 | Inconsistent dark/light theme switching                |


---

## Observations

### What deck-builder did well

- Best board review deck — zero rendering bugs, highest visual polish, consistent KPI + chart + table pattern
- Strongest content accuracy across all prompts — financial data reconciles, no arithmetic errors
- Three-layer architecture produces consistent styling (shared constants, helpers)
- Warm, intentional color palettes matched to each topic

### What MiniMax did well

- Excellent board-level content — the exec summary "3 wins / 2 concerns / 1 decision" format is textbook
- Best operational dashboard design (UAL slide 10 with progress bars)
- Consistent page number badges on every non-title slide
- Ambitious visualizations (Gantt chart, progress bars, horizontal bar overlays)

### What Anthropic did well

- Strongest writing — pasta deck copy has real personality and specific technique advice
- Highest data density in corporate prompt — charts + tables on nearly every slide
- Board review has proper financial formatting (parentheses for negatives, YoY column)
- Bold visual motifs (dark sandwich slides, decorative corner blocks, split panels)

### Improvements to make (deck-builder)

- Fix charSpacing bug that caused garbled heading on microservices slide 7
- Add chart axis labels and legends consistently
- Fill vertical space better — too much whitespace on data slides (corporate prompt exposed this)
- More chart variety needed — waterfall, sparklines, gauges
- Consider page number badges (MiniMax does this well)

### Improvements to make (all skills)

- ESG/sustainability slides are the weakest across all three skills
- Chart axis labels consistently missing units — all skills need this
- Dark-theme slides need larger minimum font for body text
- 16-slide data-dense decks stress layout engines harder than 8-12 slide decks
