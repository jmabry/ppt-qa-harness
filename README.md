# pptx-qa

A QA prompt for PPTX generation with Claude Code — shared as an example of learning how to work with agents on creative, visual tasks.

Claude can generate presentations using [Anthropic's pptx skill](https://github.com/anthropics/skills) and `pptxgenjs`. But `pptxgenjs` doesn't clip overflow — text silently renders past bounding boxes, table columns collapse, and layout bugs are invisible until you render the file.

This repo provides a `CLAUDE.md` that tells Claude to render every slide, read the images, check them against a checklist, and fix what's broken — in a loop, until clean.

## Results

| Deck | Before QA | After QA | Delta | Issues Fixed |
|------|-----------|----------|-------|-------------|
| **Corporate** (18 slides) | 17/25 | 21/25 | +4 | 16 font/layout issues |
| **Software** (6 slides) | 21/25 | 21/25 | +0 | 0 (clean baseline) |
| **Strategy** (6→7 slides) | 17/25 | 21/25 | +4 | 7 issues, slide split |
| **Total** | **55/75** | **63/75** | **+8** | 23 |

The biggest wins are on **readability** and **layout correctness** — the dimensions where "looks fine in code" diverges most from "looks fine rendered." Decks with more data density (tables, callout boxes, annotations) accumulate more rendering bugs. Full breakdown: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md)

## The cost of iterative QA

The QA loop works, but it burns wall-clock time. Our bakeoff took ~20 minutes across 4 decks — most of it in the fix-render-inspect cycle, not initial generation.

- **Rendering is serial.** LibreOffice runs single-instance; each conversion takes ~30s. No way to parallelize.
- **Each fix requires a full re-render.** Change one font size on slide 11, re-render all 18 slides.
- **Inspection is fast, fixing is slow.** Spotting issues takes seconds. Tracing font sizes through generator helpers and recalculating layout takes minutes.

Parallel subagents and targeted re-inspection (only re-reading affected slides on pass 2) helped. Two passes was enough — we never needed pass 3.

**The meta-lesson:** the best use of the QA loop is to learn what rules to add to the *generation* step so the loop has less work to do next time. Most of our 23 issues were the same bug (font too small) repeated across slides — a single generation rule ("never use fontSize below 9") would have prevented them. For creative tasks where output is visual but authoring is code, the agent needs a feedback loop that includes looking at the result. But front-loading constraints beats retrofitting fixes.

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
