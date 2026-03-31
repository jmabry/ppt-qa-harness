# Bakeoff: QA Pass

QA the baseline decks in `bakeoff/outputs/`. These were generated without QA instructions — do not regenerate them.

## Decks

| Deck | Baseline File | Baseline Generator | QA Generator (output) |
|------|---------------|--------------------|-----------------------|
| smoketest | `bakeoff/outputs/smoketest-baseline.pptx` | `bakeoff/generators/smoketest-baseline.js` | `bakeoff/generators/smoketest-qa.js` |
| corporate | `bakeoff/outputs/corporate-baseline.pptx` | `bakeoff/generators/corporate-baseline.js` | `bakeoff/generators/corporate-qa.js` |
| software | `bakeoff/outputs/software-baseline.pptx` | `bakeoff/generators/software-baseline.js` | `bakeoff/generators/software-qa.js` |
| strategy | `bakeoff/outputs/strategy-baseline.pptx` | `bakeoff/generators/strategy-baseline.js` | `bakeoff/generators/strategy-qa.js` |

## Instructions

For each deck:

1. **QA the deck.** Follow the QA process in CLAUDE.md: render slides, inspect every image against the checklist, fix the generator, re-run, iterate until clean (max 3 passes). **Do not overwrite the baseline generators.** Instead, copy each baseline generator to `bakeoff/generators/{name}-qa.js` and make all fixes there. Each time you regenerate after fixes, save the PPTX as `bakeoff/outputs/{name}-qa.pptx` (and keep `bakeoff/outputs/{name}-qa-pass{N}.pptx` snapshots). Run generators from the `bakeoff/` directory so `require('pptxgenjs')` resolves from `node_modules/`.

2. **Log results immediately.** After each deck's QA completes, append its summary to `bakeoff/outputs/RESULTS.md`:
   ```
   ## {name}
   - QA passes: {N}
   - Issues found: {N}
   - Issues fixed: {N}
   - Status: CLEAN | REMAINING
   - Remaining issues: (list, if any)
   ```

3. **Fill out the scorecard.** After all decks are processed, read the rendered slide images and `bakeoff/outputs/RESULTS.md`, then fill in `bakeoff/SCORECARD.md`. "Before QA" reflects the original baseline; "After QA" reflects the final state after fixes.

Process all decks, then fill the scorecard, then print a final summary.
