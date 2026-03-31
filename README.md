# ppt-qa-harness

QA harness for PPTX generation with Claude Code and [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

Claude generates decks using [Anthropic's pptx skill](https://github.com/anthropics/skills). This harness makes the QA loop mandatory — Claude cannot declare a deck done until a subagent has rendered and inspected every slide and returned `CLEAN`.

---

## The problem

pptxgenjs does not clip overflow. Text silently renders past its bounding box. Tables with `colW` values that don't sum correctly render column headers as vertical stacks of individual characters. You cannot see any of this without rendering the file.

The skill says to do visual QA but doesn't enforce it. Claude often stops before rendering and inspecting. The harness fixes this: `CLAUDE.md` requires a clean QA pass before Claude can declare the deck done, and `.claude/agents/pptx-qa.md` handles rendering and inspection.

---

## Setup

**1. Install LibreOffice and Poppler** (required for rendering):

```bash
brew install --cask libreoffice && brew install poppler   # macOS
sudo apt install libreoffice poppler-utils                # Debian/Ubuntu
```

**2. Install the pptx skill** (required for generation):

```bash
claude mcp add --transport stdio anthropic-skills -- npx @anthropic-ai/skills pptx
```

**3. Clone this repo and run Claude in it.** Everything is pre-wired:

- `CLAUDE.md` — instructs Claude to render slides, sanity-check, then run the QA loop
- `.claude/agents/pptx-qa.md` — the QA subagent that inspects slides and returns a bug report

Just `cd` into the repo and ask Claude to make a deck.

---

## How it works

```
User:   make a deck from brief.md
Claude: [generates deck using pptx skill]
Claude: [renders slides with soffice + pdftoppm, reads a few images]
Claude: [spawns pptx-qa agent]
Agent:  [inspects all slides in parallel, returns bug list]
Claude: [fixes generator, re-runs, re-renders, re-spawns pptx-qa]
Agent:  CLEAN — no layout issues found
Claude: Deck is at outputs/my-deck.pptx
```

The loop runs until the subagent returns `CLEAN`, capped at 3 iterations. If issues remain after 3 passes, Claude reports what was fixed and what's still outstanding.

The render step before spawning pptx-qa catches obviously broken decks (blank slides, wrong layout) before spending a full QA pass. The subagent handles deep inspection in parallel, keeping the main agent's context clean.

---

## Using in your own project

Copy these files into your project:

```
.claude/agents/pptx-qa.md    # QA subagent
CLAUDE.md                     # QA loop instruction (append to yours)
```

---

## Bakeoff

Eval comparing generation with and without the QA harness. Same skill docs, same tools, same prompts. The only variable is the `CLAUDE.md` instruction requiring a structured QA loop with a `CLEAN` gate.

Results: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md)

### Running the bakeoff

From the bakeoff directory:

```bash
cd bakeoff && claude -p PROMPT.md \
  --allowedTools "Bash Write Read Glob Grep Agent" \
  --model claude-sonnet-4-6 \
  --max-budget-usd 50
```

Existing PPTX files in `bakeoff/outputs/` are reused — only missing decks are generated. Results are written incrementally to `bakeoff/outputs/RESULTS.md`, and the scorecard is filled at the end.

To regenerate a specific deck, delete its PPTX first:

```bash
rm bakeoff/outputs/corporate-baseline.pptx
```

### Results

| Deck | Pre-QA | Post-QA | Delta | Status |
|------|--------|---------|-------|--------|
| corporate | 15/25 | 16/25 | +1 | REMAINING — systemic blank space needs layout refactor |
| software | 16/25 | 20/25 | +4 | CLEAN — all 7 issues fixed |
| strategy | 17/25 | 19/25 | +2 | REMAINING — slide 3 too content-dense for one slide |
| **Total** | **48/75** | **55/75** | **+7** | |

The harness helps most on targeted rendering bugs (software: +4). It helps less on structural problems like blank space or content density, which need the generator redesigned rather than patched.

---

## License

MIT
