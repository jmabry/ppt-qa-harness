# pptx-qa

A QA prompt for PPTX generation with Claude Code — shared as an example of learning how to work with agents on creative, visual tasks.

Claude can generate presentations using [Anthropic's pptx skill](https://github.com/anthropics/skills) and `pptxgenjs`. But `pptxgenjs` doesn't clip overflow — text silently renders past bounding boxes, table columns collapse, and layout bugs are invisible until you render the file.

This repo provides a `CLAUDE.md` that defines the QA loop (render → inspect → fix → re-run), caps passes at three, and describes an optional **parallel fix workflow** for multi-slide decks: one agent per slide (or small group), single-slide patches, then a **sequential merge** back into the generator or master XML. Full detail lives in `CLAUDE.md`.

## Results

Scores are 1–5 on five dimensions (max 25 per deck). **Delta** is rubric points gained after QA — not the same thing as how many rendering bugs we logged.

| Deck | Before QA | After QA | Delta |
|------|-----------|----------|-------|
| **Corporate** (18 slides) | 19/25 | 21/25 | +2 |
| **Software** (6 slides) | 19/25 | 21/25 | +2 |
| **Strategy** (6→7 slides) | 18/25 | 20/25 | +2 |
| **Total** | **56/75** | **62/75** | **+6** |

Separately, the QA pass found and fixed **11** concrete baseline issues in the three decks (4 + 2 + 5); strategy also split one overloaded slide so the deck grew by one slide. Itemized list: [`bakeoff/outputs/RESULTS.md`](bakeoff/outputs/RESULTS.md).

The biggest wins are on **readability** and **layout correctness** — the dimensions where "looks fine in code" diverges most from "looks fine rendered." Decks with more data density (tables, callout boxes, annotations) accumulate more rendering bugs. Dimension-by-dimension notes: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md).

## What the workflow costs (and where it helps)

The loop works, but wall-clock adds up. Our bakeoff landed around **~30 minutes** for four decks — mostly render, inspect, and implement fixes, not first-draft generation.

**Bottlenecks that stay serial**

- **Full-deck conversion** — One LibreOffice headless run per full PPTX→PDF pass (~tens of seconds per deck). After you merge parallel fixes, you still owe **one** full re-render of the whole deck for a confidence pass.
- **Merge** — After parallel patch agents finish, changes go into the canonical generator (Path A) or the unpacked master deck (Path B) in a **single merge pass** — no concurrent edits to the source of truth.

**Where parallelism pays off**

- **Inspection** — Chunk large decks (~5 slides per chunk) or split reads across subagents so you are not staring at 18 slides in one context window.
- **Fixes** — **Path A** (you have a pptxgenjs generator): one agent per slide (or 2–3 slides), each with a standalone `*-qa/patches/gen-slide-*.js` that emits a **single-slide** PPTX, a quick render to verify, then a JSON report. **Path B** (no generator, edit XML): same isolation pattern on single-slide extracts — verify on a small render, merge XML into the unpacked master afterward. Parallel agents are not fighting over the same file until merge.
- **Re-inspection on later passes** — After the first pass, only re-read slides that failed (plus neighbors for context), per `CLAUDE.md`.

**Guardrails** — Outer loop: max **3** QA iterations. Per-slide agents: max **2** fix attempts before filing partial/failed. Our bakeoff needed **one** full QA pass per deck; we never used pass 2.

**Meta-lesson** — The best ROI is still **front-loading rules into generation** (minimum font sizes, no shrink-to-fit, split tables/slides instead). Most of the **11** logged baseline issues were the same small-font class of bug repeated across slides. The QA loop proves what to codify; constraints in the generator beat chasing EMU after render.

## Setup

```bash
# Rendering tools
brew install --cask libreoffice && brew install poppler   # macOS
sudo apt install libreoffice poppler-utils                # Debian/Ubuntu

# The pptx skill
claude mcp add --transport stdio anthropic-skills -- npx @anthropic-ai/skills pptx
```

Then copy `CLAUDE.md` into your project (or append it to your existing one).

## Running the bakeoff

```bash
claude -p bakeoff/PROMPT.md \
  --allowedTools "Bash Write Read Glob Grep" \
  --model claude-opus-4-6 \
  --max-budget-usd 50
```

## License

MIT
