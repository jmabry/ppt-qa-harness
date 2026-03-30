# deck-builder-skill

A Claude Code skill for generating polished PPTX presentations with [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

## Bakeoff Results

Three skills compared on 4 identical prompts (corporate investor update, creative cooking night, software architecture migration, board strategy review). All 12 decks generated in a single pass with no QA loops — first-pass quality only.

| Skill | Corporate (UAL) | Creative (Pasta) | Software (Microservices) | Strategy (Board Review) | **Total** |
|-------|----------------|-----------------|--------------------------|-------------------------|-----------|
| **deck-builder** | 18/25 | 20/25 | 19/25 | 23/25 | **80/100** |
| Anthropic pptx | 17/25 | 21/25 | 18/25 | 22/25 | 78/100 |
| MiniMax pptx-generator | 18/25 | 18/25 | 18/25 | 22/25 | 76/100 |

**The gap is narrow (80 vs 78 vs 76). No skill dominates. Each wins in different areas.**

Scoring was done by the same agent that built deck-builder — see the [bias disclosure in the scorecard](bakeoff/SCORECARD.md#conditions).

Output decks: [`bakeoff/outputs/`](bakeoff/outputs/) — PPTX and PDF for all 12 decks, organized by prompt for side-by-side comparison.
Full scoring: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md)

## How the skills differ

### deck-builder — reliable structure, conservative visuals

This skill. Best at preventing bugs in 10+ slide decks through architectural patterns.

- **Architecture-first:** Three-layer pattern (constants → helpers → slides), Y-position chaining, config-driven templates
- **Overflow prevention:** `fitBullets`, `trimText`, `checkFit`, `estimateLines` — automated content fitting pipeline with console reporting
- **Mandatory QA loop:** LibreOffice render → subagent inspection → fix → re-render. Not optional.
- **Fewest rendering bugs** in the bakeoff (1 critical defect across 46 slides vs 4-5 for competitors)
- **Weaknesses:** Conservative layouts, underuses available slide space, limited chart variety, no page badges

Best for: **Decks where correctness matters more than visual ambition** — board reviews, status reports, anything where a text overflow or data error is worse than a plain layout.

### [Anthropic pptx](https://github.com/anthropics/skills) — best writing, highest data density

Proprietary license. Covers both generation from scratch and XML editing of existing files.

- **Strongest writing quality** — pasta deck copy has real personality; technique tips are specific and engaging
- **Highest data density** — packs charts + tables + KPI cards onto every data slide
- **XML editing workflow** — can unpack/edit/repack existing PPTX files (unique capability)
- **Bold visual motifs** — dark/light sandwich, split panels, decorative corner blocks
- **Weaknesses:** Pervasively small text (12+ of 16 slides on corporate prompt), data errors on dense decks, inconsistent theme switching

Best for: **Editing existing presentations** (only skill that covers this), or when **writing quality and data density** matter more than layout reliability.

### [MiniMax pptx-generator](https://github.com/MiniMax-AI/skills) — most visually ambitious, more breakage

MIT license. Built-in design system with 18 color palettes and 4 style recipes.

- **Most visually ambitious** — Gantt charts, progress bars, decorative shapes, dashboard layouts
- **Built-in design system** — 18 palettes, Sharp/Soft/Rounded/Pill style recipes, page number badges
- **Best space utilization** — fills slides with subtitles, additional metrics panels, callout boxes
- **Textbook exec content** — "3 wins / 2 concerns / 1 decision" board summary format
- **Weaknesses:** Most rendering defects (text overflow, wrapping bugs, illegible slides, table truncation), low-contrast body text

Best for: **Visually distinctive presentations** where you want the first draft to look designed, and you're willing to fix more rendering issues in QA.

## Known Shortcomings (deck-builder)

From the bakeoff, deck-builder's first-pass generation has recurring issues:

- **Excessive whitespace** — the biggest gap vs competitors. Data slides often fill only 60-70% of vertical space.
- **charSpacing bug** — garbled heading text on the microservices deck slide 7
- **Limited chart variety** — no waterfall, sparkline, progress bar, or gauge patterns
- **Missing chart axis labels** — charts rendered without units ($B, %, x)
- **No page number badges** — MiniMax adds these consistently
- **Conservative visual style** — reliable but not visually striking

These are fixable in the QA loop — the skill's mandatory render-inspect-fix pass catches and corrects them. But competitors produce more visually interesting first drafts.

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
  outputs/            # All 12 output decks (PPTX + PDF, organized by prompt)
    00-corporate/     # United Airlines investor update (16 slides)
    01-creative/      # Homemade pasta cooking night (8 slides)
    02-software/      # Monolith to microservices (10 slides)
    03-strategy/      # Q3 board review (12 slides)
  deck-builder/       # Generators that produced deck-builder's bakeoff decks
  prompts/            # Shared input prompts (00-03)
  SCORECARD.md        # Scores, methodology, per-slide observations, bug list
  ORCHESTRATION.md    # How to re-run the bakeoff
```

## Capability Comparison

| Capability | deck-builder | [MiniMax](https://github.com/MiniMax-AI/skills) | [Anthropic](https://github.com/anthropics/skills) |
|---|---|---|---|
| **License** | MIT | MIT | Proprietary |
| **Generation approach** | Single-file generator with helpers | Modular `slides/` dir + `compile.js` | Single-file or XML editing |
| **Architecture guidance** | Three-layer pattern, Y-chaining, config-driven templates | 7-phase workflow, theme system | None (API tutorial only) |
| **Overflow prevention** | `fitBullets`, `trimText`, `checkFit`, `estimateLines` | `fit:"shrink"` for titles | None |
| **MIN_FONT enforcement** | Yes (9pt floor, split slides) | No | No |
| **Design system** | Per-project palettes | 18 palettes, 4 style recipes, component radius | 10 palettes, typography tables |
| **Visual QA** | Mandatory render loop + subagent inspection | Content check + placeholder grep | Subagent visual inspection |
| **Template editing** | Not covered | XML unpack/edit/repack | XML unpack/edit/repack |
| **Worked examples** | 3 generators (bakeoff decks) | None included | None included |
| **i18n support** | No | Yes (Chinese fonts) | No |

## Acknowledgments

The `fit:"shrink"` tip in `pptxgenjs.md` is adapted from [MiniMax's pptx-generator skill](https://github.com/MiniMax-AI/skills/tree/main/skills/pptx-generator) (MIT licensed).

## License

MIT
