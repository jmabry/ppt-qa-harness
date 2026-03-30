# deck-builder-skill

A Claude Code skill for generating polished PPTX presentations with [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

## Bakeoff Results

Three skills compared on identical prompts. All 9 decks generated in a single pass with no QA loops — first-pass quality only.

| Skill | Creative (Pasta) | Software (Microservices) | Strategy (Board Review) | **Total** |
|-------|-----------------|--------------------------|-------------------------|-----------|
| **deck-builder** | 21/25 | 19/25 | 24/25 | **64/75** |
| Anthropic pptx | 21/25 | 18/25 | 22/25 | 61/75 |
| MiniMax pptx-generator | 18/25 | 18/25 | 22/25 | 58/75 |

Output decks: [`bakeoff/outputs/`](bakeoff/outputs/) — PPTX and PDF for all 9 decks, side by side.
Generator code: [`bakeoff/deck-builder/`](bakeoff/deck-builder/) — the JS that produced deck-builder's results.
Full scoring: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md)

## Known Shortcomings

From the bakeoff, deck-builder's first-pass generation has four recurring issues:

- **charSpacing bug** — garbled heading text on the microservices deck slide 7 (`charSpacing` applied incorrectly)
- **Missing chart legends** — charts with multiple series rendered without axis labels or legends
- **Dark background text** — body text at 10pt on dark themes is borderline too small; 11pt reads better
- **No page number badges** — MiniMax adds a small numbered badge on every non-title slide; deck-builder doesn't

These are fixable in the QA loop — the skill's mandatory render-inspect-fix pass is where they get caught and corrected.

## Decision Guide

**Use deck-builder** if you need architecture patterns that scale (10+ slides, config-driven templates, overflow prevention) and want a mandatory QA loop.

**Use MiniMax pptx-generator** if you want a built-in design system (18 color palettes, 4 style recipes) and don't need architectural guidance. Strong board-level content format ("3 wins / 2 concerns / 1 decision").

**Use Anthropic pptx** if you need to edit existing PPTX files (XML unpack/edit/repack workflow) or want the strongest writing quality on first pass.

## Why this exists

Most AI-assisted slide generation takes the wrong path: convert content to markdown, pipe it through a Python library, and hope the output looks professional. This fails in practice:

- **Markdown destroys intent.** Headings, bullets, and paragraphs are the only primitives. You lose layout control — columns, cards, callout boxes, positioned graphics — the moment you flatten to markdown.
- **Python PPTX libraries are layout-blind.** python-pptx gives you XML manipulation, not visual design. You're computing EMU offsets and hoping for the best.
- **Iteration is slow.** Generate, open PowerPoint, squint at the result, edit code, repeat. No fast feedback loop.

This skill takes a different approach:

- **JavaScript-first generation** with pptxgenjs — coordinates in inches, a clean API, and a Node.js runtime that's fast to iterate with.
- **Architecture patterns that scale** — constants, helpers, and config-driven templates prevent the layout bugs that plague decks with 10+ slides.
- **Mandatory visual QA** — LibreOffice renders the PPTX to images so you can actually see what you built before calling it done.

## Install

### Claude Code (CLI)

```bash
# From your project directory
claude skill add jmabry/deck-builder-skill
```

### Manual

Copy `SKILL.md`, `architecture.md`, and `pptxgenjs.md` into your project's `.claude/skills/deck-builder/` directory.

## Dependencies

```bash
npm install pptxgenjs
pip install "markitdown[pptx]"
brew install --cask libreoffice   # macOS — required for visual QA
```

## Repo Structure

```
SKILL.md              # Skill entry point (Claude Code reads this first)
architecture.md       # Three-layer pattern, Y-chaining, content fitting
pptxgenjs.md          # API reference, pitfalls, helper patterns
bakeoff/
  outputs/            # All 9 output decks (PPTX + PDF, checked in)
  deck-builder/       # Generators that produced deck-builder's bakeoff decks
    01-creative/      # Homemade pasta — cards, two-column layouts
    02-software/      # Monolith to microservices — diagrams, timeline bar
    03-strategy/      # Q3 board review — KPI cards, charts, tables
  prompts/            # Shared input prompts (01-03)
  SCORECARD.md        # Scores, per-slide observations, bug list
  ORCHESTRATION.md    # How to re-run the bakeoff
```

## Examples

The `bakeoff/deck-builder/` generators are the canonical code reference — they produced the output decks above and demonstrate the architecture patterns on realistic prompts:

- **`01-creative/gen-pasta-deck.js`** — warm palette, card templates, two-column layouts
- **`02-software/gen-microservices-deck.js`** — dark theme, architecture diagram, phased timeline bar
- **`03-strategy/gen-board-deck.js`** — KPI card grid, bar/line charts, exec summary tables

Run any of them:

```bash
cd bakeoff/deck-builder/01-creative
npm install pptxgenjs
node gen-pasta-deck.js
```

Each follows the three-layer architecture from `architecture.md`: constants layer, helper functions that return next-Y for chaining, config-driven templates for repeated layouts.

## How this compares

| Capability | deck-builder-skill | [MiniMax pptx-generator](https://github.com/MiniMax-AI/skills) | Anthropic pptx |
|---|---|---|---|
| **License** | MIT | MIT | Proprietary |
| **Generation approach** | Single-file generator with helpers | Modular `slides/` dir + `compile.js` | Single-file or XML editing |
| **Architecture guidance** | Three-layer pattern (constants, helpers, slides), Y-chaining, config-driven templates | 7-phase workflow, theme system | None (API tutorial only) |
| **Overflow prevention** | `fitBullets`, `trimText`, `checkFit`, `estimateLines` | `fit:"shrink"` for titles | "Be careful with text" |
| **MIN_FONT enforcement** | Yes (9pt floor, split slides) | No | No |
| **Content fitting pipeline** | Yes (automated with console reporting) | No | No |
| **Design system** | Per-project (bakeoff generators show different palettes) | 18 built-in color palettes, 4 style recipes, component radius system | 10 color palettes, typography tables |
| **Visual QA** | Mandatory LibreOffice render loop, subagent inspection, cross-slide consistency checks | markitdown content check, placeholder grep | Subagent visual inspection |
| **Template editing** | Not covered (generation only) | XML unpack/edit/repack workflow | XML unpack/edit/repack workflow |
| **Worked examples** | 3 generators (pasta, microservices, board review) | None included | None included |
| **Pitfalls documented** | 8 (with JS fundamentals context) | 8+ (including async gotcha) | 8 |
| **i18n support** | No | Yes (Chinese font handling) | No |

## Acknowledgments

The `fit:"shrink"` tip in `pptxgenjs.md` is adapted from [MiniMax's pptx-generator skill](https://github.com/MiniMax-AI/skills/tree/main/skills/pptx-generator) (MIT licensed).

## License

MIT
