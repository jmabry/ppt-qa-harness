---
name: deck-builder
description: "Use this skill to create polished PPTX presentations from scratch using pptxgenjs. Covers architecture patterns for decks with 10+ slides, automated content fitting, build-time overflow detection, and a mandatory visual QA loop. Trigger when creating slide decks, pitch decks, business reviews, or any multi-slide presentation that needs to look professional."
---

# Deck Builder Skill

**Recommended model: Opus.** Visual QA (reading rendered slide screenshots to catch overflow, misalignment, clipped text) requires strong image understanding.

## Quick Reference

| Task | Guide |
|------|-------|
| Understand the approach | Read [architecture.md](architecture.md) first |
| pptxgenjs API + patterns | Read [pptxgenjs.md](pptxgenjs.md) |
| See working examples | Browse `examples/` |

---

## Creating Presentations

**Read [architecture.md](architecture.md) before writing any code.** It covers the three-layer pattern (constants, helpers, slides) that prevents layout bugs at scale.

**Read [pptxgenjs.md](pptxgenjs.md) for API details** — text, shapes, images, tables, charts, and common pitfalls.

**Generation and QA are one atomic operation.** After every `node gen-*.js` run, you MUST render with LibreOffice, read the slide screenshots, and fix any issues before reporting to the user. Never generate without QA. Never QA without fixing. The loop is: generate, render, inspect, fix, re-render, confirm clean.

---

## Reading Existing PPTX

```bash
# Text extraction
pip install "markitdown[pptx]"
python -m markitdown presentation.pptx
```

---

## Visual QA Loop

**Treat every render as broken until proven otherwise.** pptxgenjs does not clip overflow, so text silently bleeds past bounding boxes. The only way to verify layout is to render the PPTX and look at the actual pixels.

### Content Check

Extract text from the generated file and scan for gaps:

```bash
python -m markitdown output.pptx
```

Verify: all sections present, correct slide order, no placeholder text left behind, no truncated sentences.

### Render-and-Read Loop

You cannot visually QA a PPTX without rendering it. Use LibreOffice to convert to images:

```bash
soffice --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

This produces `slide-01.jpg`, `slide-02.jpg`, etc. After fixing a specific slide, re-render just that page:

```bash
pdftoppm -jpeg -r 150 -f 3 -l 3 output.pdf slide-fixed
```

Install LibreOffice if missing:
```bash
brew install --cask libreoffice   # macOS
sudo apt install libreoffice      # Ubuntu/Debian
```

### What to Look For

Read every slide image and check:

- Text that overflows its bounding box (pptxgenjs renders it, just past the boundary)
- Any element whose bottom edge crosses below the footer line
- Sections or cards that feel crammed together without whitespace
- Misaligned columns, uneven card grids, or ragged vertical edges
- Body text smaller than MIN_FONT — if you can't read it at arm's length, it's too small
- Poor readability from low contrast (test: squint at the slide — can you still parse it?)

**Delegate inspection to a subagent.** After writing the generator code, your eyes are biased toward seeing what you intended. A fresh subagent reading the screenshots will catch issues you'll miss. Send the subagent the slide images with a brief description of what each slide should contain, and ask it to report every problem it finds.

### Cross-Slide Consistency

Individual slides can look fine in isolation but feel inconsistent as a deck. After checking each slide:

- Flip through adjacent pairs — do fonts, colors, and margins stay consistent?
- Pick 2-3 random non-adjacent slides and compare styling
- Confirm all body text meets the MIN_FONT floor
- Verify section labels, card borders, and spacing use the same constants throughout

### Fix-and-Verify Discipline

Every issue you find requires a fix, and every fix requires re-rendering. Don't batch fixes and hope they all worked — one adjustment often shifts something else.

The cycle: find issue, fix code, regenerate, re-render the affected slide, confirm it's clean, move to the next issue. Only report success after a complete pass with zero new findings.

---

## Dependencies

- `npm install pptxgenjs` — slide generation
- `pip install "markitdown[pptx]"` — text extraction
- LibreOffice (`soffice`) — PDF conversion (**required**, not optional)
- Poppler (`pdftoppm`) — PDF to images
