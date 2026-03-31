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
| Visual appeal      | TBD             | TBD          |       |
| Content accuracy   | TBD             | TBD          |       |
| Layout correctness | TBD             | TBD          |       |
| Data density       | TBD             | TBD          |       |
| Readability        | TBD             | TBD          |       |
| **Total**          | **TBD/25**      | **TBD/25**   |       |

## Software (Monolith to Microservices)

6-slide technical presentation for engineering leadership.

| Dimension          | Without harness | With harness | Notes |
| ------------------ | --------------- | ------------ | ----- |
| Visual appeal      | TBD             | TBD          |       |
| Technical accuracy | TBD             | TBD          |       |
| Layout correctness | TBD             | TBD          |       |
| Data visualization | TBD             | TBD          |       |
| Readability        | TBD             | TBD          |       |
| **Total**          | **TBD/25**      | **TBD/25**   |       |

## Strategy (Q3 Board Review)

6-slide board-level presentation for Series B SaaS company.

| Dimension          | Without harness | With harness | Notes |
| ------------------ | --------------- | ------------ | ----- |
| Visual appeal      | TBD             | TBD          |       |
| Executive presence | TBD             | TBD          |       |
| Layout correctness | TBD             | TBD          |       |
| Data density       | TBD             | TBD          |       |
| Readability        | TBD             | TBD          |       |
| **Total**          | **TBD/25**      | **TBD/25**   |       |

---

## Overall

| Prompt        | Without harness | With harness | Delta |
| ------------- | --------------- | ------------ | ----- |
| **corporate** | TBD/25          | TBD/25       | TBD   |
| **software**  | TBD/25          | TBD/25       | TBD   |
| **strategy**  | TBD/25          | TBD/25       | TBD   |
| **Total**     | **TBD/75**      | **TBD/75**   | **TBD** |

---

## Bugs Found

### Without harness

*(Fill after running `./run-bakeoff.sh generate` and rendering)*

| Deck | Slide(s) | Issue | Severity |
| ---- | -------- | ----- | -------- |
| — | — | — | — |

### With harness

*(Fill after running `./run-bakeoff.sh harness` and rendering)*

| Deck | Slide(s) | Issue | Notes |
| ---- | -------- | ----- | ----- |
| — | — | — | — |

---

## What the harness caught

*(Fill after both runs — which bugs were present without harness but fixed with harness)*

| Deck | Slide(s) | Bug | Fixed by harness? |
| ---- | -------- | --- | ----------------- |
| — | — | — | — |
