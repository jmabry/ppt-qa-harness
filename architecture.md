# Architecture Overview

How to structure a pptxgenjs deck generator that scales to 15+ slides without layout bugs.

## The problem

Raw pptxgenjs is imperative: you call `addText`, `addShape`, `addTable` with absolute x/y/w/h coordinates. This works for 1-2 slides. At 15 slides it collapses:

- **Magic numbers everywhere** — `y: 2.34` means nothing a week later
- **Cascading breakage** — change one element's height, everything below shifts
- **Copy-paste drift** — slide 7's header is 2px different from slide 4's
- **Overflow invisible in code** — text silently renders past its box

## The solution: three layers

```
+---------------------------------------------+
|  Slide definitions (data + layout calls)    |  <- what each slide contains
+---------------------------------------------+
|  Helpers (addHeader, addSectionLabel, etc.)  |  <- reusable building blocks
+---------------------------------------------+
|  Constants + utilities (PAD, BODY_TOP, etc.) |  <- single source of truth
+---------------------------------------------+
```

### Layer 1: Constants and utilities

Define every measurement once. Derive dependent values.

```javascript
const W = 10, H = 5.625;           // 16:9 slide dimensions
const PAD = 0.5;                    // left/right margin
const TITLE_H = 0.5;               // title text height
const BODY_TOP = TITLE_H + 0.12;   // derived: first content Y
const BODY_W = W - PAD * 2;        // derived: usable width
const FOOTER_Y = 5.35;             // footer line position
const CONTENT_BOTTOM = FOOTER_Y - 0.12;  // hard lower boundary
const SECTION_GAP = 0.12;          // minimum whitespace between sections
const MIN_FONT = 9;               // hard floor — no body text below this
```

**Key rules:**
- `CONTENT_BOTTOM` is the floor. No element's `y + h` should exceed it.
- `MIN_FONT` is the readability floor. Split slides rather than shrink font below 9pt.

Utilities at this layer handle text fitting:

```javascript
trimText(text, maxChars)       // shorten to fit — rewrite as complete thought, never leave ellipsis
fitBullets(items, max, chars)  // cap + truncate, log drops to console
checkFit(label, text, ...)     // warn at build time if text won't fit box
```

### Layer 2: Helpers

Small functions that encapsulate repeated patterns. The critical design rule: **helpers that consume vertical space return the next Y position.**

```javascript
function addSectionLabel(slide, text, y, opts) {
  // render label + underline
  return y + 0.24;  // <- next element goes here
}
```

This enables Y-chaining:

```javascript
let y = BODY_TOP;
y = addSectionLabel(s, "Section A", y);
addBullets(s, items, PAD, y, BODY_W, 0.6);
y += 0.6 + SECTION_GAP;
y = addSectionLabel(s, "Section B", y);  // automatically positioned
```

Typical helper set:

| Helper | Purpose | Returns |
|--------|---------|---------|
| `addHeader(slide, title)` | Title + brand tag | void |
| `addFooter(slide)` | Page number + accent bar | void |
| `addSectionLabel(slide, text, y, opts)` | Label + underline | **next Y** |
| `addBullets(slide, items, x, y, w, h, fs)` | Bulleted text block (enforces MIN_FONT) | void |
| `addSubHeader(slide, text, x, y, w, opts)` | Bold caps sub-section label | **next Y** |
| `addCardRow(slide, cards, x, y, w, h, cols)` | Grid of cards with header + body | void |
| `addCommitmentCard(slide, title, desc, x, y, w, h)` | Card with colored top accent | void |
| `addChipLabel(slide, text, x, y, opts)` | Colored chip badge | void |
| `twoColumnLayout(gap)` | Standard two-column dimensions | `{leftX, rightX, colW}` |
| `estimateLines(text, fontSize, boxW)` | Wrapping-aware line count | number |

### Layer 3: Slide definitions

Each slide is either:

**A. Inline block** — for one-off layouts (cover, architecture diagram):
```javascript
{
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, "Solution Architecture");
  addFooter(s);
  // ... custom layout code using helpers
}
```

**B. Config-driven template** — for repeated layouts (status slides):
```javascript
function addStatusSlide(cfg) {
  const s = pres.addSlide({ masterName: "MASTER" });
  addHeader(s, cfg.title);
  addFooter(s);

  // Badge row, left column, right sidebar — all parameterized by cfg
  // All Y positions chained, all text auto-fitted
}

// Each slide is just data:
addStatusSlide({ title: "Feature X", priority: "P0", status: [...], blockers: [...] });
addStatusSlide({ title: "Feature Y", priority: "P1", status: [...], blockers: [...] });
```

Six slides, zero layout bugs, because the template is correct once.

## Patterns that work

### Y-position chaining
Never hardcode a Y coordinate that depends on content above it. Always derive from the previous element's position + height.

### Content-aware height allocation
Estimate how many lines each section needs, divide available height proportionally:
```javascript
const totalLines = outcomeLines + statusLines + blockerLines;
const lineH = usableH / totalLines;
const outcomeH = outcomeLines * lineH;
const statusH = statusLines * lineH;
```

### Build-time overflow detection
Before placing text in a fixed-height box, estimate whether it fits:
```javascript
checkFit("Mission box", missionText, 8.5, boxW, boxH, 1.15);
// Console: warning OVERFLOW: "Mission box" ~0.52" exceeds box 0.41"
```
Fix the warning before rendering. This catches what visual QA would catch, but earlier.

### Console reporting (not in-deck appendix)
When `fitBullets` drops bullets or `trimText` truncates, the change is logged to console. The agent reads the console output and proactively summarizes changes to the user. **If cuts change the deck's focus or key takeaways, ask the user for guidance with numbered options before proceeding.** Use plan mode for significant editorial decisions.

### Content prioritization
When source content exceeds slide capacity, prioritize what the audience needs to act on:
- **Status reports**: blockers > risks > next steps > accomplishments
- **Pitch decks**: value proposition > differentiation > proof points > team
- General rule: if the audience can only read one section, which one drives a decision?

### Cover slide helper
Cover slides are always custom, but still benefit from a helper to avoid Y-position overlap:
```javascript
function addCoverSlide(cfg) {
  const s = pres.addSlide();
  s.background = { color: cfg.bgColor };
  let y = 1.0;
  s.addText(cfg.title, { x: PAD, y, w: 6, h: 0.8, fontSize: 36, bold: true, ... });
  y += 0.9;
  s.addText(cfg.subtitle, { x: PAD, y, w: 6, h: 0.4, fontSize: 16, ... });
  y += 0.5;
  s.addText(cfg.meta, { x: PAD, y, w: 6, h: 0.25, fontSize: 11, ... });
}
```
Even on cover slides, chain Y positions to prevent overlap when titles wrap.

### Empty section handling
When a section has no content:
- **Expected sections** (e.g., blockers on a status slide): show "No blockers" — the absence is informative
- **Optional sections** (e.g., next quarter goals): omit and reallocate vertical space to remaining sections
- **Uncertain**: use plan mode to ask the user whether to include or omit

### Slide master for shared background
```javascript
pres.defineSlideMaster({
  title: "MASTER",
  background: { color: "ffffff" },
  objects: []  // footer/header added by helpers, not master
});
```
Keep the master minimal. Helpers give you more control than master objects.

## Patterns that fail

### Fixed-proportion heights
```javascript
// BAD — breaks when content varies between slides
const statusH = availH * 0.35;
```
Use content-aware allocation instead.

### Hardcoded Y offsets
```javascript
// BAD — breaks when anything above changes
const section2Y = 2.34;
```
Chain through return values.

### Guessing rowH for tables
```javascript
// BAD — overflow renders past the cell, not clipped
{ rowH: 0.38 }
```
Remove `rowH` for variable content. Let pptxgenjs auto-size.

### Hand-editing content to fit
```javascript
// BAD — future agent won't know the rules
status: ["Shortened bullet that lost meaning"]
```
Use `fitBullets`/`trimText` in the generator. The original stays intact; the console log records what was cut.

## Style consistency

Helpers enforce consistent styling by centralizing font sizes, colors, margins, and spacing. But helpers can't catch everything — visual QA must also compare slides against each other, not just check each slide in isolation.

During QA, compare adjacent slides and spot-check 2-3 random non-adjacent pairs:
- Font sizes consistent across similar elements (all >= MIN_FONT)
- Section labels same style throughout deck
- Card/box borders, fills, and radii consistent
- Margins and padding use constants (PAD), not ad-hoc values
- No jarring style change from one slide to the next

When content doesn't fit at MIN_FONT, **split the slide** rather than shrinking the font. One dense slide at 7pt is worse than two readable slides at 9pt.

## Build order

1. **Constants** — measurements, colors, design tokens, MIN_FONT
2. **Utilities** — `trimText`, `fitBullets`, `checkFit`, `estimateLines`
3. **Helpers** — `addHeader`, `addFooter`, `addSectionLabel`, `addBullets`, `addSubHeader`, `addCardRow`, `addCoverSlide`
4. **Templates** — `addStatusSlide(cfg)` and similar config-driven functions
5. **Slides** — one block per slide, using helpers and templates
6. **`writeFile`** — single output call
