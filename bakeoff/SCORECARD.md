# Bakeoff Scorecard

## Conditions

Same prompts, same pptx skill, same tools. Two conditions:

| Condition | What's different |
|-----------|-----------------|
| **Before QA** | Baseline generation — skill docs only, no QA instructions |
| **After QA** | `CLAUDE.md` present — Claude must render, inspect, and fix every slide |

The baseline is not "no QA capability" — Claude has bash and LibreOffice available either way. It's "discretionary QA" vs "mandatory QA." A win here means the prompt helps even when the tools are already there.

## Scoring

Each deck is rated 1–5 on 5 dimensions (3 universal + 2 prompt-specific). Scores are based on visual inspection of rendered slide images.

| Score | Meaning |
|-------|---------|
| **5** | Excellent — polished, would present as-is |
| **4** | Good — minor issues that don't hurt usability |
| **3** | Adequate — noticeable issues but functional |
| **2** | Below average — significant issues hurt usability |
| **1** | Poor — broken rendering, missing content, unusable |

### Universal Dimensions

| Dimension | 1 (Poor) | 5 (Excellent) |
|-----------|----------|---------------|
| **Visual appeal** | Plain text on white, no design effort | Cohesive palette, varied layouts, professional polish |
| **Layout correctness** | Text overflow, clipping, overlapping elements | No overflow, consistent margins, proper whitespace |
| **Readability** | Tiny fonts, low contrast, cramped spacing | Clear hierarchy, good contrast, comfortable font sizes |

### Prompt-Specific Dimensions

| Prompt | Dimension 1 | Dimension 2 |
|--------|-------------|-------------|
| **corporate** (UAL) | **Content accuracy** — real financials, plausible data | **Data density** — charts, tables, KPI cards |
| **software** | **Technical accuracy** — correct terminology, sound architecture | **Data visualization** — diagrams, charts, flow representation |
| **strategy** | **Executive presence** — would you show this to a board? | **Data density** — charts, tables, KPI cards |

---

## Corporate (United Airlines Investor Update)

18-slide institutional investor deck with specific financials, fleet data, and strategic metrics.

| Dimension | Before QA | After QA | Notes |
|-----------|-----------|----------|-------|
| Visual appeal | 4 | 4 | Consistent navy/gold palette, professional layouts. QA didn't change design. |
| Content accuracy | 4 | 4 | Dense, detailed financial data presented logically. Unchanged by QA. |
| Layout correctness | 3 | 4 | Before: Latin callout clipped (slide 7), missing chart axis (slide 15). After: all resolved. |
| Data density | 5 | 5 | Very data-rich — tables, charts, KPI cards, narrative callouts throughout. |
| Readability | 3 | 4 | Before: slide 7 footnote 8.5pt, slide 15 chart uninterpretable without y-axis. After: all resolved. |
| **Total** | **19/25** | **21/25** | |

## Software (Monolith to Microservices)

6-slide technical presentation for engineering leadership.

| Dimension | Before QA | After QA | Notes |
|-----------|-----------|----------|-------|
| Visual appeal | 4 | 4 | Clean dark theme, well-structured tables. No changes needed. |
| Technical accuracy | 5 | 5 | Correct strangler fig pattern, sound ADR, realistic migration plan. |
| Layout correctness | 3 | 4 | Before: Risk Matrix Mitigation column clipped (slide 5). After: resolved. |
| Data visualization | 4 | 4 | Risk matrix, observability stack, DORA metrics — good for 6 slides. |
| Readability | 3 | 4 | Before: footer dark-bg text ~7–8pt on slide 5. After: bumped to 11pt, clearly legible. |
| **Total** | **19/25** | **21/25** | |

## Strategy (Q3 Board Review)

6-slide board-level presentation for Series B SaaS company.

| Dimension | Before QA | After QA | Notes |
|-----------|-----------|----------|-------|
| Visual appeal | 4 | 4 | Polished dark/light panel design. Unchanged. |
| Executive presence | 4 | 4 | Good narrative arc; clear board asks on slide 6. Unchanged. |
| Layout correctness | 3 | 4 | Before: table clipping and overflow on slides 3–4. After: all resolved. |
| Data density | 4 | 4 | Rich data throughout; slides 3–4 dense but functional. |
| Readability | 3 | 4 | Before: body text 7–8pt on slides 3–4. After: all at 9pt minimum; chart labels legible. |
| **Total** | **18/25** | **20/25** | |

---

## Overall

| Prompt | Before QA | After QA | Delta |
|--------|-----------|----------|-------|
| **corporate** | 19/25 | 21/25 | +2 |
| **software** | 19/25 | 21/25 | +2 |
| **strategy** | 18/25 | 20/25 | +2 |
| **Total** | **56/75** | **62/75** | **+6** |

---

## Issues

### Before QA (baseline generation)

| Deck | Issues | Summary |
|------|--------|---------|
| corporate | 4 | Latin callout panel overflow (slide 7); A321XLR footnote 8.5pt (slide 7); missing y-axis on Emissions Intensity chart (slide 15); ~30% blank space on slide 15. |
| software | 2 | Risk Matrix Mitigation column clipped (slide 5); footer dark-bg text ~7–8pt (slide 5). |
| strategy | 5 | Sub-9pt text in Churn Deep Dive and Key Hires tables (slides 3–4); At-Risk Accounts Risk Signal column clipped (slide 3); bar chart labels 7–7.5pt (slide 4); attrition note 7.5pt (slide 4). |
| smoketest | 0 | Clean — single slide, no issues. |

### After QA

| Deck | Fixed | Remaining | Summary |
|------|-------|-----------|---------|
| corporate | 4 | 0 | All issues resolved in 1 pass. |
| software | 2 | 0 | Both issues resolved in 1 pass. |
| strategy | 5 | 0 | All issues resolved in 1 pass. |
| smoketest | 0 | 0 | No fixes needed. |

## Timing

- Total wall-clock time: ~30 minutes
- QA passes used: 1 of 3 max (all decks)
- Rendering: LibreOffice single-instance, ~30s per deck per pass
- Inspection parallelized via 4 subagents simultaneously
- Fix implementation: 3 parallel fix agents, then 3 parallel merge agents
