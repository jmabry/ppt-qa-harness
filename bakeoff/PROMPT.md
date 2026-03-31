# Bakeoff: Baseline vs QA Harness

Generate PPTX decks for each prompt in `prompts/`, then QA and fix them.

## Setup

pptxgenjs is pre-installed in `node_modules/`. Use `require('pptxgenjs')`.

## Prompts

Generate one deck per prompt file in `prompts/`:

| Prompt | File | Output |
|--------|------|--------|
| smoketest | `prompts/smoketest.md` | `outputs/smoketest-baseline.pptx` |
| corporate | `prompts/corporate.md` | `outputs/corporate-baseline.pptx` |
| software | `prompts/software.md` | `outputs/software-baseline.pptx` |
| strategy | `prompts/strategy.md` | `outputs/strategy-baseline.pptx` |

## Instructions

For each prompt:

1. **Check for existing output.** If `outputs/{name}-baseline.pptx` already exists, skip generation and go straight to step 3.

2. **Generate the deck.** Read the prompt file, write a Node.js generator to `generators/{name}-baseline.js`, and run it to produce the PPTX in `outputs/`.

3. **QA the deck.** Follow the QA process in CLAUDE.md: render slides, sanity-check, spawn pptx-qa, fix, iterate until CLEAN (max 3 passes). Each time you regenerate after fixes, copy the previous PPTX to `outputs/{name}-baseline-pass{N}.pptx` before overwriting.

4. **Log results immediately.** After each deck's QA completes (or is skipped), append its summary to `outputs/RESULTS.md` right away — do not wait until all decks are done:
   ```
   ## {name}
   - QA passes: {N}
   - Issues found: {N}
   - Issues fixed: {N}
   - Status: CLEAN | REMAINING
   - Remaining issues: (list, if any)
   ```

5. **Fill out the scorecard.** After all decks are processed, read the rendered slide images and `outputs/RESULTS.md`, then fill in `SCORECARD.md` — score each deck 1-5 on each dimension, write notes, fill the bugs tables, and compute totals. The "Without harness" column reflects the pre-QA state (pass1); "With harness" reflects the final post-QA state.

Process all prompts, then fill the scorecard, then print a final summary.
