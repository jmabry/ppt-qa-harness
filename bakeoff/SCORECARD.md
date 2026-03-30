# Bakeoff Scorecard

## Conditions

All 9 decks were generated in a single pass using `claude --print` with each skill's docs injected as system prompt. **No QA loops ran** — subagents were denied Bash permission, so no skill got to render, inspect, or fix its output. Scores reflect first-pass generation quality only.

**Scoring bias disclosure:** The deck-builder skill is this repo's own skill, and scoring was done by the same agent that built it. To mitigate this: (1) visual inspections were delegated to fresh subagents with no knowledge of which skill produced which deck, (2) specific slide screenshots are referenced so readers can verify scores against the actual outputs in `bakeoff/outputs/`, and (3) where my initial scores favored deck-builder, I re-examined by comparing equivalent slides side-by-side and adjusted.

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

| Prompt                              | Dimension 1                                                                   | Dimension 2                                                       |
| ----------------------------------- | ----------------------------------------------------------------------------- | ----------------------------------------------------------------- |
| **corporate** (United Airlines)  | **Content accuracy** — real financials, plausible data, internal consistency   | **Data density** — charts, tables, KPI cards, metrics per slide  |
| **software** (Microservices)     | **Technical accuracy** — correct terminology, sound architecture              | **Data visualization** — diagrams, charts, flow representation  |
| **strategy** (Board Review)      | **Executive presence** — would you show this to a board?                      | **Data density** — charts, tables, KPI cards, metrics             |

---

## Corporate (United Airlines Investor Update) ★

This is the most representative real-world test — a data-dense 16-slide corporate investor deck generated from a long-form research document with specific financials, fleet data, and strategic metrics.

| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                            |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 3            | 4         | 4         | DB has consistent navy/gold but excessive whitespace on slides 6, 8, 10, 13. MiniMax fills space better with subtitles, progress bars, and decorative shapes. Anthropic's dark/light sandwich is bold. |
| Content accuracy   | 5            | 4         | 3         | DB data is internally consistent with no errors. MiniMax has a leverage chart/KPI mismatch (3x vs 2.6x). Anthropic has a MAX 8 arithmetic error (215 vs 255) and Q3 EPS shows 5.51 instead of 3.51. |
| Layout correctness | 3            | 3         | 3         | DB wastes vertical space on 5+ slides. MiniMax has text wrapping bugs (EWR split, % signs on slide 10) and illegible ESG slide. Anthropic has pervasively small text across 12+ slides. |
| Data density       | 3            | 4         | 4         | DB is chart-light — only 6 of 16 slides have charts, with a repetitive KPI+bullets pattern. MiniMax uses KPI cards, progress bars, multi-year charts, and additional metrics tables. Anthropic packs charts + tables densely. |
| Readability        | 4            | 3         | 3         | DB text is readable where present but chart labels lack units. MiniMax slide 15 is nearly illegible, slides 5/13 have low-contrast footnotes. Anthropic has pervasively small sub-text and footnotes across most slides. |
| **Total**          | **18/25**    | **18/25** | **17/25** |                                                                                                                                                                  |

## Software (Monolith to Microservices)

| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                                                    |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 4            | 3         | 4         | DB dark navy+teal is polished. Anthropic dark theme with colored accents is professional. MiniMax navy+gold is functional but monotone.                                                  |
| Technical accuracy | 5            | 5         | 5         | All three correctly apply strangler fig, DORA metrics, service mesh, event-driven architecture.                                                                                          |
| Layout correctness | 3            | 3         | 3         | DB has garbled text on slide 7 (corrupted heading). MiniMax has header text wrapping ("LIKELIHOO/D") and Gantt chart clipping. Anthropic has text truncation on slide 6 ("Foundatio n"). |
| Data visualization | 4            | 3         | 3         | DB has a clean architecture diagram, phased timeline bar, and org chart. MiniMax has a Gantt chart (ambitious but clipped). Anthropic architecture diagram is basic; no actual charts.   |
| Readability        | 3            | 4         | 3         | DB dark theme has small text + garbled heading. MiniMax yellow-on-navy is high contrast. Anthropic dark theme has small text and truncation.                                             |
| **Total**          | **19/25**    | **18/25** | **18/25** |                                                                                                                                                                                          |

## Strategy (Q3 Board Review)

| Dimension          | deck-builder | MiniMax   | Anthropic | Notes                                                                                                                                                |
| ------------------ | ------------ | --------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| Visual appeal      | 4            | 4         | 4         | All three are professional and board-appropriate. MiniMax's dark sidebar and traffic-light exec summary are textbook. DB and Anthropic are clean but not clearly ahead. |
| Executive presence | 5            | 5         | 5         | All three produce board-ready decks with exec summary, specific asks, confidential markings, and financial appendices.                               |
| Layout correctness | 5            | 4         | 4         | DB has zero rendering bugs. MiniMax slide 12 has a truncated cash flow table. Anthropic slides 5+9 have charts without legends.                      |
| Data density       | 5            | 5         | 5         | All three pack KPI cards, bar/line charts, tables, pipeline breakdowns, and full P&L appendices. Excellent across the board.                         |
| Readability        | 4            | 4         | 4         | All three are dense but appropriate for board materials. Small text in some tables/chart legends across all skills.                                  |
| **Total**          | **23/25**    | **22/25** | **22/25** |                                                                                                                                                      |

---

## Overall

| Skill            | corporate | software | strategy | Grand Total |
| ---------------- | --------- | -------- | -------- | ----------- |
| **deck-builder** | 18/25        | 19/25       | 23/25       | **60/75**   |
| **MiniMax**      | 18/25        | 18/25       | 22/25       | **58/75**   |
| **Anthropic**    | 17/25        | 18/25       | 22/25       | **57/75**   |

**Takeaway:** The gap is narrow (60 vs 58 vs 57). Each skill wins in different areas. No skill dominates.

---

## Skill Comparison

### deck-builder — reliable structure, weak density

**Strengths:** Fewest rendering bugs overall. The three-layer architecture (constants → helpers → slides) produces consistent styling and internally consistent data. Best content accuracy — no arithmetic errors across any prompt. Y-position chaining prevents overlap. Board review deck was the cleanest of any skill on any prompt.

**Weaknesses:** Underuses available slide space — excessive whitespace on data slides, particularly the corporate prompt where 5+ slides have 30-40% dead space at the bottom. Repetitive layout pattern (KPI cards + bullets appears on too many slides). Chart variety is limited — no waterfall, sparkline, gauge, or progress bar patterns. Least visually ambitious of the three.

### MiniMax — ambitious visuals, more rendering defects

**Strengths:** Most visually ambitious — attempts Gantt charts, progress bars, decorative shapes, dark-background accent slides. Best operational dashboards (UAL slide 10 with progress bars). Page number badges on every non-title slide. Fills slide space well with subtitles, additional metrics tables, and callout boxes. Textbook exec summary format ("3 wins / 2 concerns / 1 decision").

**Weaknesses:** Most rendering defects — text overflow (pasta slide 4), text wrapping bugs (UAL slides 8, 10), header truncation (microservices "LIKELIHOO/D"), illegible ESG slide, table truncation. The ambition that produces better layouts also produces more breakage. Low-contrast body text on multiple pasta slides.

### Anthropic — best writing, small text problem

**Strengths:** Strongest writing quality — pasta deck copy has real personality and specific technique tips. Highest data density on corporate and board review prompts. Bold visual motifs (dark sandwich slides, split panels, decorative corner blocks). Proper financial formatting (parentheses for negatives, YoY columns). Best slide variety on the pasta prompt.

**Weaknesses:** Pervasively small text — on the corporate prompt, 12+ of 16 slides have sub-text, footnotes, or table cells that would be illegible when projected. Data errors on the corporate prompt (MAX 8 arithmetic, Q3 EPS value). Inconsistent dark/light theme switching feels accidental rather than structured. Unicode emoji icons on strategy slides look unprofessional.

---

## Key Issues Found

### Critical (rendering bugs)

| Skill        | Deck          | Slide | Issue                                                                 |
| ------------ | ------------- | ----- | --------------------------------------------------------------------- |
| deck-builder | Microservices | 7     | Garbled/corrupted heading text ("A_N_E_R_T_H_R_G_E_S...")             |
| MiniMax      | Microservices | 6     | Gantt chart "Monolith Retire" clipped at right edge                   |
| MiniMax      | UAL           | 8, 10 | EWR airport code split across lines; % signs wrap to next line        |
| MiniMax      | UAL           | 15    | ESG slide text nearly illegible (faded/small bullets in all 3 cards)  |
| Anthropic    | Microservices | 6     | Timeline headings truncated mid-word ("Foundatio n", "Accelerat e")   |
| Anthropic    | UAL           | 7     | MAX 8 arithmetic error (215 orders but 142+113=255)                   |
| Anthropic    | UAL           | 3     | Q3 EPS shows $5.51 instead of $3.51                                   |

### Moderate

| Skill        | Deck          | Slide    | Issue                                                          |
| ------------ | ------------- | -------- | -------------------------------------------------------------- |
| deck-builder | UAL           | multiple | Excessive whitespace — content fills only 60-70% of slide      |
| deck-builder | UAL           | 5, 7     | Chart axis labels missing units ($B, %)                        |
| MiniMax      | Microservices | 5        | "LIKELIHOOD" header wraps to "LIKELIHOO/D"                     |
| MiniMax      | Board Review  | 12       | Cash Flow table truncated at bottom                            |
| MiniMax      | UAL           | 12       | Leverage chart vs KPI card data inconsistency (3x vs 2.6x)    |
| Anthropic    | Board Review  | 5, 9     | Charts with multiple series but no legend                      |
| Anthropic    | UAL           | multiple | Pervasively small sub-text and footnotes (12+ slides)          |
| Anthropic    | UAL           | 6, 9, 13 | Inconsistent dark/light theme switching                        |

---

## Improvements to make (deck-builder)

- **Fill vertical space** — the biggest gap vs competitors. Add subtitles, additional metrics panels, or expand chart areas to use available slide real estate.
- **More chart variety** — add waterfall, sparkline, progress bar, and gauge patterns to the pptxgenjs guide.
- **Fix charSpacing bug** that caused garbled heading on microservices slide 7.
- **Chart axis labels** — always include units ($B, %, x) on axes and data labels.
- **Page number badges** — MiniMax does this consistently and it adds polish.
- **Bolder visual elements** — the skill produces reliable but conservative output. The architecture docs could encourage more decorative shapes, dark accent slides, and varied card designs.

## Improvements to make (all skills)

- ESG/sustainability slides are the weakest across all three skills.
- Chart axis labels consistently missing units — all skills need this.
- Dark-theme slides need larger minimum font for body text.
- 16-slide data-dense decks stress layout engines harder than 8-12 slide decks — all skills showed more defects on the corporate prompt.
