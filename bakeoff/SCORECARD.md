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

16–18 slide institutional investor deck with specific financials, fleet data, and strategic metrics.

| Dimension | Before QA | After QA | Notes |
|-----------|-----------|----------|-------|
| Visual appeal | 4 | 4 | Strong navy/gold palette, good visual hierarchy. QA didn't change design. |
| Content accuracy | 4 | 4 | Plausible financials throughout. Unchanged by QA. |
| Layout correctness | 2 | 4 | Before: 12 critical font violations, text overflow on slides 11/16. After: all resolved. |
| Data density | 5 | 5 | Exceptional — tables, bar charts, callout boxes, KPI strips on every slide. |
| Readability | 2 | 4 | Before: pervasive 7-8pt text in callouts/annotations across 10+ slides. After: all text meets minimums. |
| **Total** | **17/25** | **21/25** | |

## Software (Monolith to Microservices)

6-slide technical presentation for engineering leadership.

| Dimension | Before QA | After QA | Notes |
|-----------|-----------|----------|-------|
| Visual appeal | 4 | 4 | Clean dark theme, well-structured tables. No changes needed. |
| Technical accuracy | 5 | 5 | Correct strangler fig pattern, sound ADR, realistic migration plan. |
| Layout correctness | 4 | 4 | No issues found on any slide. Clean baseline. |
| Data visualization | 4 | 4 | Risk matrix, before/after metrics, bar chart. Good for 6 slides. |
| Readability | 4 | 4 | All fonts at or above minimums already. |
| **Total** | **21/25** | **21/25** | |

## Strategy (Q3 Board Review)

6-slide (now 7 after split) board-level presentation for Series B SaaS company.

| Dimension | Before QA | After QA | Notes |
|-----------|-----------|----------|-------|
| Visual appeal | 4 | 4 | Polished dark theme with color-coded KPIs. Unchanged. |
| Executive presence | 3 | 4 | Before: slide 3 was overwhelmingly dense — not board-ready. After: split into two focused slides. |
| Layout correctness | 2 | 4 | Before: slide 3 crammed 5+ sections into one slide, text near-illegible. Slide 5 clipping. After: clean. |
| Data density | 5 | 5 | Charts, tables, KPI cards throughout. Split added a slide but kept all content. |
| Readability | 3 | 4 | Before: slide 3 text well below 9pt, dark bg text below 11pt. After: all minimums met. |
| **Total** | **17/25** | **21/25** | |

---

## Overall

| Prompt | Before QA | After QA | Delta |
|--------|-----------|----------|-------|
| **corporate** | 17/25 | 21/25 | +4 |
| **software** | 21/25 | 21/25 | +0 |
| **strategy** | 17/25 | 21/25 | +4 |
| **Total** | **55/75** | **63/75** | **+8** |

---

## Issues

### Before QA (baseline generation)

| Deck | Issues | Summary |
|------|--------|---------|
| corporate | 16 | Systemic undersized text (7-8.5pt) in callout boxes, annotations, and appendix tables across slides 6-18. Text overflow on slides 11, 16. |
| software | 0 | Clean baseline — all slides passed inspection. |
| strategy | 7 | Slide 3 critically overloaded (2 critical); font sizes below minimums on dark backgrounds (5 minor). |
| smoketest | 0 | Clean — single slide, no issues. |

### After QA

| Deck | Fixed | Remaining | Summary |
|------|-------|-----------|---------|
| corporate | 16 | 0 | 47 font edits in pass 1; 3 layout fixes in pass 2. All clean. |
| software | 0 | 0 | No fixes needed. |
| strategy | 7 | 0 | Slide split + font fixes in pass 1; position adjustments in pass 2. All clean. |
| smoketest | 0 | 0 | No fixes needed. |

## Timing

- Total wall-clock time: ~20 minutes
- QA passes used: 2 of 3 max (corporate and strategy); 1 of 3 (smoketest and software)
- Rendering bottleneck: LibreOffice (`soffice`) runs single-instance, ~30s per deck per pass
- Inspection parallelized via subagents (3 decks simultaneously)
- Fix implementation: subagents for pass 1 bulk edits, manual for pass 2 targeted fixes
