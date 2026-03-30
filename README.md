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

## License

MIT
