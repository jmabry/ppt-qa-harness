# ppt-qa-harness

QA harness for PPTX generation with Claude Code and [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

**Use [Anthropic's pptx skill](https://github.com/anthropics/skills) to generate. Use this harness to make the QA loop mandatory.**

---

## The problem

pptxgenjs does not clip overflow. Text silently renders past its bounding box. A table with `colW` values that don't sum correctly renders column headers as vertical stacks of individual characters. You cannot see any of this without rendering the file.

The skill says to do visual QA but doesn't enforce it. Even when LibreOffice is available, Claude often stops before rendering and inspecting. The harness fixes this: `CLAUDE.md` requires a clean QA pass before Claude can declare the deck done, and `.claude/agents/pptx-qa.md` handles the rendering and inspection.

---

## Setup

**1. Install LibreOffice and Poppler** (required for rendering):

```bash
brew install --cask libreoffice && brew install poppler   # macOS
sudo apt install libreoffice poppler-utils                # Debian/Ubuntu
```

**2. Clone this repo and run Claude in it.** Everything is pre-wired:

- `CLAUDE.md` — tells Claude to run the QA loop after generating any PPTX
- `.claude/agents/pptx-qa.md` — the QA subagent that renders and inspects slides
- `.claude/settings.json` — PostToolUse hook that auto-renders `.pptx` files

No manual file copying required. Just `cd` into the repo and run `claude`.

---

## How it works

```
User:   make a deck from brief.md
Claude: [generates deck using Anthropic pptx skill]
Claude: [spawns pptx-qa agent]
Agent:  [renders slides, finds issues, returns bug report]
Claude: [fixes generator, re-runs, re-spawns pptx-qa]
Agent:  CLEAN — no layout issues found
Claude: Deck is at outputs/my-deck.pptx
```

The `CLAUDE.md` instruction prevents Claude from stopping early. The `pptx-qa` subagent handles rendering and inspection, so the main agent's context stays clean. The loop runs until the subagent finds nothing.

**Simpler alternative:** `.claude/settings.json` contains a PostToolUse hook that auto-renders any `.pptx` on disk and injects the slide image paths into Claude's context. No subagent required — Claude inspects the images directly.

---

## Using in your own project

Copy these files into your project:

```
.claude/agents/pptx-qa.md    # QA subagent
CLAUDE.md                     # QA loop instruction (append to yours)
```

Optionally copy `.claude/settings.json` for the auto-render hook.

---

## Best practices for pptxgenjs

These patterns prevent the bugs the `pptx-qa` agent is designed to catch.

### Layout

**Always validate colW sums.** `colW` values that don't add up to `w` cause columns to collapse to near-zero width, rendering text one character per line vertically.

```javascript
const colW = [2.5, 1.5, 1.5, 1.5];
console.assert(Math.abs(colW.reduce((a,b)=>a+b,0) - w) < 0.01, `colW mismatch on ${label}`);
```

**Never use charSpacing on text objects.** It corrupts rendering in LibreOffice — headers become garbled character strings.

**Use Y-chaining for anything that stacks vertically.** Hard-coded `y` values cascade: change one element's height and everything below shifts. Return the next Y from every helper:

```javascript
let y = BODY_TOP;
y = addSectionLabel(slide, "Section A", y);  // returns y + label height
y = addTable(slide, rows, y);                 // starts where label ended
```

**Set CONTENT_BOTTOM and never exceed it.** `FOOTER_Y - 0.12` is the hard floor. pptxgenjs renders past it silently — LibreOffice shows the overflow.

### Typography

**MIN_FONT = 9pt.** Below 9pt is illegible when projected. Split the slide rather than shrink further. Dark backgrounds need 11pt minimum.

**Always set chart legends and axis labels explicitly.** pptxgenjs doesn't add them by default — `showLegend: true`, `catAxisLabelFontSize`, `valAxisLabelFontSize`.

### Structure

**Three-layer pattern for decks with 10+ slides:**

```
Constants + utilities  →  single source of truth for measurements and colors
Helpers                →  reusable building blocks, Y-chaining
Slide definitions      →  data + layout calls only, no magic numbers
```

---

## Bakeoff

We ran Anthropic's pptx skill on 3 professional prompts twice — same skill docs, same tools (bash + LibreOffice), same prompts. The only difference: the harness adds a CLAUDE.md instruction requiring a structured `pptx-qa` agent loop with a `CLEAN` gate before declaring done. Without it, the skill docs still encourage QA but don't enforce it — Claude may or may not inspect its output before stopping.

Source prompts, generators, and rendered slide images: [`bakeoff/`](bakeoff/)

Full before/after results: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md)

Run the eval: `./bakeoff/run-bakeoff.sh` (defaults to full pipeline, see script header for config docs)

---

## Lessons learned

We built pptxgenjs generation patterns iteratively on real decks — board reviews, investor updates, architecture presentations. Each pattern came from a specific failure: a table grew by one row and pushed everything off the slide; bullet text overflowed silently; elements rendered into the footer invisibly.

Those are real patterns. But the main lesson wasn't a new generation technique — it was that Claude stops before inspecting its output. The QA loop runs when you require it to. That's the harness.

---

## License

MIT
