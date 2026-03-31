# Bakeoff Scorecard

## Conditions

Two runs of Anthropic's pptx skill on the same prompts:

| Run | Tools | QA behavior |
|-----|-------|-------------|
| **Without harness** | bash + LibreOffice available | Skill docs encourage QA but don't enforce it — Claude may inspect its output or may not |
| **With harness** | bash + LibreOffice available | `CLAUDE.md` mandates a `pptx-qa` agent loop (`.claude/agents/pptx-qa.md`); Claude cannot declare done until the agent returns `CLEAN` |

Both runs use the same skill docs, tools, and prompts. The only variable is the harness CLAUDE.md instruction. This means the without-harness run is not a "no QA capability" baseline — it's a "discretionary QA" baseline. A harness win under these conditions is stronger: the structured loop helps even when Claude already knows it should inspect. Scores are based on visual inspection of rendered slide images (LibreOffice PDF export, 150 DPI JPEG).

## Scoring Methodology

Each deck is rated 1-5 on 5 dimensions (3 universal + 2 prompt-specific).

| Score | Meaning                                              |
| ----- | ---------------------------------------------------- |
| **5** | Excellent — polished, no issues, would present as-is |
| **4** | Good — minor issues that don't hurt usability        |
| **3** | Adequate — noticeable issues but still functional    |
| **2** | Below average — significant issues hurt usability    |
| **1** | Poor — broken rendering, missing content, unusable   |

### Universal Dimensions

| Dimension              | 1 (Poor)                                      | 5 (Excellent)                                          |
| ---------------------- | --------------------------------------------- | ------------------------------------------------------ |
| **Visual appeal**      | Plain text on white, no design effort         | Cohesive palette, varied layouts, professional polish  |
| **Layout correctness** | Text overflow, clipping, overlapping elements | No overflow, consistent margins, proper whitespace     |
| **Readability**        | Tiny fonts, low contrast, cramped spacing     | Clear hierarchy, good contrast, comfortable font sizes |

### Prompt-Specific Dimensions

| Prompt              | Dimension 1                                                                 | Dimension 2                                                      |
| ------------------- | --------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| **corporate** (UAL) | **Content accuracy** — real financials, plausible data, internal consistency | **Data density** — charts, tables, KPI cards, metrics per slide |
| **software**        | **Technical accuracy** — correct terminology, sound architecture            | **Data visualization** — diagrams, charts, flow representation  |
| **strategy**        | **Executive presence** — would you show this to a board?                    | **Data density** — charts, tables, KPI cards, metrics            |

---

## Corporate (United Airlines Investor Update)

16-18 slide institutional investor deck with specific financials, fleet data, and strategic metrics.

| Dimension          | Without harness | With harness | Notes |
| ------------------ | --------------- | ------------ | ----- |
| Visual appeal      | 3               | 3            | Cohesive navy/gold palette, good title slide with KPI cards. Blank lower-thirds on many slides persist. |
| Content accuracy   | 4               | 4            | Real UAL financials, plausible TRASM/CASM breakdowns, fleet data, CapEx timeline. Internal consistency strong. |
| Layout correctness | 2               | 3            | Pre-QA: 8 issues. 5 fixed (targeted rendering bugs). Systemic blank space on 9+ slides remains — needs full layout refactor. |
| Data density       | 3               | 3            | Tables and KPI cards present but blank lower halves waste space. Appendix slide (18) is well-packed. |
| Readability        | 3               | 3            | Fonts adequate. Slide 14 title wraps to two lines, chart label clipped. Blank space not illegible, just wasteful. |
| **Total**          | **15/25**       | **16/25**    |       |

## Software (Monolith to Microservices)

6-slide technical presentation for engineering leadership.

| Dimension          | Without harness | With harness | Notes |
| ------------------ | --------------- | ------------ | ----- |
| Visual appeal      | 4               | 4            | Dark tech theme with teal/cyan accents. Consistent across all 6 slides. Professional look. |
| Technical accuracy | 4               | 4            | DORA metrics correctly cited, strangler fig rationale sound, squad ownership model realistic. |
| Layout correctness | 2               | 4            | Pre-QA: 7 issues. All 7 fixed across 3 passes — CLEAN. Biggest improvement in the bakeoff. |
| Data visualization | 3               | 4            | Tables render correctly post-fix. Color-coded legends, rejected/chosen badges on architecture decision slide. |
| Readability        | 3               | 4            | Post-fix: clear hierarchy, comfortable font sizes, good contrast on dark background. |
| **Total**          | **16/25**       | **20/25**    |       |

## Strategy (Q3 Board Review)

6-slide board-level presentation for Series B SaaS company.

| Dimension          | Without harness | With harness | Notes |
| ------------------ | --------------- | ------------ | ----- |
| Visual appeal      | 4               | 4            | Dark executive theme, strong use of red/green/orange for status. Board-ready color palette. |
| Executive presence | 4               | 4            | Decision-driven: approve/discuss badges, key metrics front-and-center, board asks with financials. |
| Layout correctness | 2               | 3            | Pre-QA: 5 issues. 3 fixed. Remaining: column clipping on slide 2, sub-9pt text and colW mismatch on slide 3, footnote overflow slide 4. |
| Data density       | 5               | 5            | Extremely dense — KPIs, pipeline, churn, GTM, customer health all packed in 6 slides. |
| Readability        | 2               | 3            | Slide 3 still too content-dense with sub-9pt text. Improved elsewhere. Should split slide 3. |
| **Total**          | **17/25**       | **19/25**    |       |

---

## Overall

| Prompt        | Without harness | With harness | Delta |
| ------------- | --------------- | ------------ | ----- |
| **corporate** | 15/25           | 16/25        | +1    |
| **software**  | 16/25           | 20/25        | +4    |
| **strategy**  | 17/25           | 19/25        | +2    |
| **Total**     | **48/75**       | **55/75**    | **+7** |

---

## Bugs Found

### Without harness (pre-QA state)

Issues present in original generation before QA loop ran.

| Deck | Issues | Summary |
| ---- | ------ | ------- |
| corporate | 8 | Blank lower-thirds on most slides, title overflow slide 14, chart label clip slide 14 |
| software | 7 | Various layout/rendering issues across 6 slides |
| strategy | 5 | Column clipping, sub-9pt text, footnote overflow |
| smoketest | 0 | Clean from the start |

### With harness (post-QA state)

Remaining issues after max 3 QA passes.

| Deck | Remaining | Summary |
| ---- | --------- | ------- |
| corporate | 11 | Systemic blank space (9+ slides) — needs full layout refactor. Title wrap + chart clip on slide 14. |
| software | 0 | CLEAN — all 7 issues fixed in 3 passes |
| strategy | 4 | Slide 2 "Trend" column clipped, slide 3 sub-9pt text + colW mismatch, slide 4 footnote truncated |
| smoketest | 0 | CLEAN — no issues to begin with |

---

## What the harness caught and fixed

| Deck | Issues found | Issues fixed | Key fixes |
| ---- | ------------ | ------------ | --------- |
| corporate | 8 | 5 | Targeted rendering bugs fixed; systemic blank space correctly identified but too structural for iterative fixes |
| software | 7 | 7 | All layout/rendering issues resolved — tables, overflow, clipping all corrected |
| strategy | 5 | 3 | Partial improvement; remaining issues are content-density problems (slide 3 needs to split) |
| smoketest | 0 | 0 | N/A |
