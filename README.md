# deck-builder-skill

A Claude Code skill for generating polished PPTX presentations with [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

## Why this exists

Most AI-assisted slide generation takes the wrong path: convert content to markdown, pipe it through a Python library, and hope the output looks professional. This fails in practice:

- **Markdown destroys intent.** Headings, bullets, and paragraphs are the only primitives. You lose layout control — columns, cards, callout boxes, positioned graphics — the moment you flatten to markdown.
- **Python PPTX libraries are layout-blind.** python-pptx gives you XML manipulation, not visual design. You're computing EMU offsets and hoping for the best.
- **Iteration is slow.** Generate, open PowerPoint, squint at the result, edit code, repeat. No fast feedback loop.

This skill takes a different approach:

- **JavaScript-first generation** with pptxgenjs — coordinates in inches, a clean API, and a Node.js runtime that's fast to iterate with.
- **HTML for early prototyping** — when you need to see content layout quickly, an HTML preview renders in seconds. Final output is always PPTX, but the inner loop can be much faster.
- **Architecture patterns that scale** — constants, helpers, and config-driven templates prevent the layout bugs that plague decks with 10+ slides.
- **Mandatory visual QA** — LibreOffice renders the PPTX to images so you can actually see what you built before calling it done.

## What it does

Teaches Claude how to create multi-slide presentation decks that don't break at scale. Covers:

- **Three-layer architecture** — constants, helpers, slides — that prevents magic numbers, cascading breakage, and copy-paste drift
- **Automated content fitting** — `fitBullets`, `trimText`, `checkFit` for overflow prevention (pptxgenjs does not clip text)
- **Config-driven templates** — one function, many slides, zero drift
- **MIN_FONT enforcement** — never shrink below 9pt; split slides instead
- **Visual QA loop** — LibreOffice rendering catches what code review can't

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
examples/             # Complete working generators
  mtb-trails/         # Mountain biking trail guide
  cub-scouts/         # Cub Scout campout planner
  cooking-class/      # Brunch cooking class
bakeoff/              # Skill comparison harness (see bakeoff/ORCHESTRATION.md)
```

## Examples

See `examples/` for complete working generators that demonstrate every pattern:

- **`mtb-trails/`** — Mountain biking trail guide (8 slides, config-driven trail cards, two-column layouts)
- **`cub-scouts/`** — Cub Scout campout planner (8 slides, config-driven activity template, tables)
- **`cooking-class/`** — Brunch cooking class (8 slides, config-driven recipe template, badge system)

Run any example:

```bash
cd examples/mtb-trails
npm install pptxgenjs
node gen-mtb-deck.js
# Output: output/mtb-trail-guide.pptx
```

Each example demonstrates the architecture patterns from `architecture.md`: constants layer, helper functions that return next-Y for chaining, config-driven templates for repeated layouts, and automated content fitting.

## How this compares

There are other PPTX skills for AI coding agents. Here's how they differ:

| Capability | deck-builder-skill | [MiniMax pptx-generator](https://github.com/MiniMax-AI/skills) | Anthropic pptx |
|---|---|---|---|
| **License** | MIT | MIT | Proprietary |
| **Generation approach** | Single-file generator with helpers | Modular `slides/` dir + `compile.js` | Single-file or XML editing |
| **Architecture guidance** | Three-layer pattern (constants, helpers, slides), Y-chaining, config-driven templates | 7-phase workflow, theme system | None (API tutorial only) |
| **Overflow prevention** | `fitBullets`, `trimText`, `checkFit`, `estimateLines` | `fit:"shrink"` for titles | "Be careful with text" |
| **MIN_FONT enforcement** | Yes (9pt floor, split slides) | No | No |
| **Content fitting pipeline** | Yes (automated with console reporting) | No | No |
| **Design system** | Per-project (examples show different palettes) | 18 built-in color palettes, 4 style recipes, component radius system | 10 color palettes, typography tables |
| **Slide type catalog** | By example (architecture.md patterns) | 5 page types with detailed layouts | Layout suggestions in Design Ideas |
| **Visual QA** | Mandatory LibreOffice render loop, subagent inspection, cross-slide consistency checks | markitdown content check, placeholder grep | Subagent visual inspection |
| **Template editing** | Not covered (generation only) | XML unpack/edit/repack workflow | XML unpack/edit/repack workflow |
| **Worked examples** | 3 complete generators (trail guide, campout, cooking class) | None included | None included |
| **Pitfalls documented** | 8 (with JS fundamentals context) | 8+ (including async gotcha) | 8 |
| **i18n support** | No | Yes (Chinese font handling) | No |

Each skill optimizes for different things. This one focuses on **architecture patterns and QA discipline** for decks that need to scale beyond 10 slides without layout bugs. MiniMax focuses on **design systems and modular compilation**. Anthropic covers both **generation and XML editing** but without architectural guidance.

### Bakeoff

The `bakeoff/` directory contains a harness for generating decks from identical prompts using all three skills and comparing outputs. See [`bakeoff/ORCHESTRATION.md`](bakeoff/ORCHESTRATION.md) for setup and usage.

## Acknowledgments

The `fit:"shrink"` tip in `pptxgenjs.md` is adapted from [MiniMax's pptx-generator skill](https://github.com/MiniMax-AI/skills/tree/main/skills/pptx-generator) (MIT licensed).

## License

MIT
