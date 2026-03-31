# pptx-qa

QA prompt for PPTX generation with Claude Code and pptxgenjs.

## Generating presentations

Use [Anthropic's pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) for generation. Install it via Claude Code:

```
claude mcp add --transport stdio anthropic-skills -- npx @anthropic-ai/skills pptx
```

### Generation rules

- **Minimum font sizes:** Body text at least 9pt. Text on dark backgrounds at least 11pt. Never shrink to fit — split across slides instead.
- **Table cell text:** At least 8pt. If a table has too many columns to fit, split it or transpose.

## QA process

Apply this process after generating a PPTX file, or when QA-ing an existing one.

1. **Render all slides to images:**
   ```bash
   SLIDE_DIR="outputs/<name>-slides"
   mkdir -p "$SLIDE_DIR"
   soffice --headless --convert-to pdf <file.pptx> --outdir "$SLIDE_DIR"
   pdftoppm -jpeg -r 120 "$SLIDE_DIR/<name>.pdf" "$SLIDE_DIR/slide"
   rm "$SLIDE_DIR/<name>.pdf"
   ```

2. **Read and inspect every slide image.** Check for:
   - Text overflow past bounding boxes or slide edges
   - Table columns clipped at the right edge
   - Vertical character stacks (letters stacking one-per-line in columns)
   - Content below the footer line
   - Blank space (>30% empty lower half with no content reason)
   - Font size below minimums (body < 9pt; dark background < 11pt)
   - Garbled headers (letters out of order, mid-word breaks)
   - Missing chart elements (axes without labels, charts without needed legends)

   For large decks (15+ slides), chunk the slides into groups of ~5 and inspect one group at a time to manage context.

3. **Fix every issue found**, re-run the generator, re-render, and re-inspect the affected slides.

4. **Repeat until clean — maximum 3 iterations.** On re-inspection, only re-read slides that had issues (plus their immediate neighbors for context).

5. Only declare the deck done when all slides pass inspection, or report what's still outstanding after 3 iterations.

## Dependencies

LibreOffice (`soffice`) and Poppler (`pdftoppm`) must be installed for rendering:

```bash
brew install --cask libreoffice && brew install poppler   # macOS
sudo apt install libreoffice poppler-utils                # Debian/Ubuntu
```

## Bakeoff

The `bakeoff/` directory contains an eval comparing generation with and without this QA prompt. See `bakeoff/PROMPT.md` for details.
