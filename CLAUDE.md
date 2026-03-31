# ppt-qa-harness

QA harness for PPTX generation with Claude Code and pptxgenjs.

## Generating presentations

Use [Anthropic's pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) for generation. Install it via Claude Code:

```
claude mcp add --transport stdio anthropic-skills -- npx @anthropic-ai/skills pptx
```

Or clone it manually and pass as a system prompt — see `bakeoff/run-bakeoff.sh` for an example.

## QA harness

The `pptx-qa` agent **only renders, inspects slides, and returns a bug list** (or `CLEAN`). It does not modify code or regenerate files—you do.

After generating a PPTX file:

1. Spawn the `pptx-qa` agent with the output file path
2. For every issue it reports, fix the generator and re-run
3. Re-spawn `pptx-qa` on the new output
4. Repeat until `pptx-qa` returns `CLEAN` — **maximum 3 QA iterations**
5. Only then tell the user the deck is done

If after 3 iterations issues remain, report what was fixed and what's still outstanding — do not loop further. Do not report success until you have a clean QA pass or have exhausted the iteration limit.

## Dependencies

LibreOffice (`soffice`) and Poppler (`pdftoppm`) must be installed for rendering:

```bash
brew install --cask libreoffice && brew install poppler   # macOS
sudo apt install libreoffice poppler-utils                # Debian/Ubuntu
```

## Bakeoff

The `bakeoff/` directory contains an automated eval comparing generation with and without the harness. Run `./bakeoff/run-bakeoff.sh` for details.
