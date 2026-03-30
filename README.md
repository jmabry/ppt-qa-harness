# deck-builder-skill

A Claude Code skill for generating polished PPTX presentations with [pptxgenjs](https://github.com/gitbrent/PptxGenJS).

> **Honest caveat:** This skill documents real learnings from iterating on actual decks — the architecture patterns, overflow helpers, and QA loop all came from hitting real bugs in production. But learnings from one context don't automatically translate to a better general-purpose skill. The bakeoff shows deck-builder finishing last on first-pass quality against both competitors. The patterns help *when the QA loop runs*. If you're evaluating which skill to use, read the [bakeoff results](#bakeoff-results) before deciding.

## Bakeoff Results

Three skills compared on 3 identical prompts — all data-heavy professional decks (corporate investor update, software architecture migration, board strategy review). This is the hardest category for layout correctness: multi-column tables, dense KPI grids, and 16+ slides are exactly where layout bugs compound. All 9 decks generated in a single pass with no QA loops — first-pass quality only. Scored by visual inspection of LibreOffice-rendered slide images, blind to which skill produced each deck.

| Skill | Corporate (UAL) | Software (Microservices) | Strategy (Board Review) | **Total** |
|-------|----------------|--------------------------|-------------------------|-----------|
| **Anthropic pptx** | 21/25 | 23/25 | 22/25 | **66/75** |
| MiniMax pptx-generator | 15/25 | 16/25 | 20/25 | **51/75** |
| **deck-builder** | 15/25 | 15/25 | 10/25 | **40/75** |

**Anthropic wins. deck-builder came last** — systematic layout failures across all three prompts (vertical table overflow on 4 strategy slides, font corruption on software slide 5, 40-50% blank space throughout the corporate deck). All are caught by the mandatory QA loop, but this bakeoff measured first-pass only. **Without QA, deck-builder is the least reliable of the three. With QA, it produces the fewest residual bugs.**

Scoring blind to skill identity — see the [bias disclosure and full scoring in the scorecard](bakeoff/SCORECARD.md#conditions).

Output decks: [`bakeoff/outputs/`](bakeoff/outputs/) — PPTX and PDF for all 9 decks.
Full scoring: [`bakeoff/SCORECARD.md`](bakeoff/SCORECARD.md)

## How the skills differ

### [Anthropic pptx](https://github.com/anthropics/skills) — best first-pass output, highest data density

Proprietary license. Covers both generation from scratch and XML editing of existing files.

- **Best first-pass quality** — won all three prompts in the bakeoff (66/75 total)
- **Highest data density** — 18-slide UAL deck with two full appendix slides, all packed with tables and KPI cards
- **Bold visual motifs** — dark/light sandwich, left-border accent bars, colored severity badges
- **Strong writing** — analytical narrative framing on context boxes, specific numbers throughout
- **XML editing workflow** — can unpack/edit/repack existing PPTX files (unique capability)
- **Weaknesses:** Minor right-edge column truncation on competitive/roadmap tables; some footnotes too small

Best for: **First-pass quality without QA**, editing existing presentations, or when data density and visual polish matter most.

### [MiniMax pptx-generator](https://github.com/MiniMax-AI/skills) — most visually ambitious, inconsistent layout

MIT license. Built-in design system with 18 color palettes and 4 style recipes.

- **Most visually ambitious** — progress bars, page badges, decorative shapes, dashboard layouts
- **Built-in design system** — 18 palettes, Sharp/Soft/Rounded/Pill style recipes
- **Best architecture diagrams** — hierarchical service diagram vs simple box chains in other skills
- **Richer data coverage** — includes merge conflict rate, rollback rate, Phase confidence percentages
- **Weaknesses:** Systematic right-column clipping on multi-year tables; title slide text stack collision; slides 4-5 broken in software deck

Best for: **Visually distinctive presentations** when you're willing to fix column-width clipping issues in QA.

### deck-builder — architecture patterns, QA-dependent

This skill. Best at preventing bugs through mandatory render-inspect-fix loops — but worst first-pass quality of the three.

- **Architecture-first:** Three-layer pattern (constants → helpers → slides), Y-position chaining, config-driven templates
- **Overflow prevention:** `fitBullets`, `trimText`, `checkFit`, `estimateLines` — automated content fitting pipeline
- **Mandatory QA loop:** LibreOffice render → subagent inspection → fix → re-render. Not optional.
- **Weaknesses (first-pass):** Vertical table overflow bug (strategy slides 2-5 unreadable), font corruption on section labels (software slide 5), 40-50% blank space on 8+ corporate slides, arithmetic error in fuel sensitivity callout

Best for: **Final-quality decks where the QA loop runs** — the mandatory render-inspect-fix pass catches all of the above. Without QA, this skill produces the least reliable output.

## Known Shortcomings (deck-builder)

From the bakeoff, deck-builder's first-pass generation has critical bugs:

- **Vertical table overflow** — when `colW` array doesn't sum to `w`, columns collapse to near-zero and text renders vertically (one character per line). Fix: assert `sum(colW) ≈ w` before every table call.
- **charSpacing font corruption** — applying `charSpacing` to any text object produces garbled output ("OBERSMEBILITY STACK", "CRITIC AL"). Remove from all section labels.
- **Excessive whitespace** — 40-50% blank lower halves on 8+ slides in the corporate deck. Add vertical fill check.
- **Arithmetic errors** — fuel sensitivity stated as $40M (should be ~$400M). Verify calculations.
- **Conservative visual style** — flat compared to Anthropic's motif-driven approach

All rendering bugs are caught by the mandatory QA loop. The skill is only as good as its QA discipline.

## Why this exists

This skill was built iteratively on real decks — board reviews, investor updates, technical architecture presentations. Each pattern in it came from hitting a real bug: Y-chaining came from cascading layout breaks when a table grew by one row, `fitBullets` came from text silently overflowing its box in LibreOffice, the QA loop came from realizing you can't see any of this without rendering.

That iteration produced genuine learnings. It didn't necessarily produce a better general-purpose skill.

The bakeoff is the honest test: all three skills got the same prompts, the same constraints, no QA. deck-builder came last. The architecture patterns didn't prevent a `colW` summation bug that made 4 of 6 strategy slides unreadable, or a charSpacing call that garbled a section label. A skill built by a team with more design investment (Anthropic) or a richer built-in component library (MiniMax) produced better first-pass output. The patterns help — but only when the QA loop runs, and only for the failure modes they were designed to catch.

**What the skill is actually good for:**

- **Decks where correctness matters more than first-pass looks** — board reviews, data-dense investor updates, anything where a broken table or arithmetic error is worse than a plain layout
- **Long decks (10+ slides)** — the three-layer pattern and Y-chaining compound over many slides; short decks don't stress the system enough to show the benefit
- **Understanding what goes wrong** — the bakeoff generators and bugs are documented, so you can see exactly what breaks and why
- **Running the full QA loop** — the mandatory render-inspect-fix pass is where this skill earns its keep; skip it and you'd be better off with Anthropic's skill

## Install

### Claude Code (CLI)

```bash
# From your project directory
claude skill add jmabry/deck-builder-skill
```

### Manual

Copy the `skill/` directory into your project as `.claude/skills/deck-builder/`.

## Dependencies

```bash
npm install pptxgenjs
pip install "markitdown[pptx]"
brew install --cask libreoffice   # macOS — required for visual QA
```

## Repo Structure

```
skill/                # Install this directory as .claude/skills/deck-builder/
  SKILL.md            # Skill entry point (Claude Code reads this first)
  architecture.md     # Three-layer pattern, Y-chaining, content fitting
  pptxgenjs.md        # API reference, pitfalls, helper patterns
bakeoff/
  outputs/            # All 9 output decks (PPTX + rendered slide images)
                      # {prompt}-{skill}.pptx  (corporate, software, strategy × 3 skills)
  generators/         # Generator scripts ({prompt}-{skill}.js)
  prompts/            # Shared input prompts (corporate.md, software.md, strategy.md)
  SCORECARD.md        # Scores, methodology, per-slide observations, bug list
  run-bakeoff.sh      # Harness to re-run the bakeoff
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
| **Worked examples** | 9 generators (3 prompts × 3 skills) | None included | None included |
| **i18n support** | No | Yes (Chinese fonts) | No |

## Acknowledgments

The `fit:"shrink"` tip in `pptxgenjs.md` is adapted from [MiniMax's pptx-generator skill](https://github.com/MiniMax-AI/skills/tree/main/skills/pptx-generator) (MIT licensed).

## License

MIT
