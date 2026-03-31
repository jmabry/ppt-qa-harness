# pptx-qa

QA prompt for PPTX generation with Claude Code and pptxgenjs.

## Generating presentations

Use [Anthropic's pptx skill](https://github.com/anthropics/skills/tree/main/skills/pptx) for generation. Install it via Claude Code:

```
/plugin install document-skills@anthropic-agent-skills
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

### Subagent rules for QA fixes

When delegating QA fixes to subagents, follow these patterns to avoid repeated failures:

- **One agent per slide (or per 2–3 slides).** Don't send an agent a full deck of fixes — context dilution causes it to miss or silently drop issues.
- **Require re-inspection in the prompt.** Explicitly tell the agent: "After fixing, read the re-rendered slide image and describe what you see." Agents skip this if not required.
- **Ask for before/after values.** "Report the original EMU value and what you changed it to." This makes fixes auditable and catches edits to the wrong element.
- **Require structured output.** For each slide the agent touches, it must report: (1) what issue it tried to fix, (2) what values it changed (before → after), (3) what the re-rendered image shows.
- **Ask for failures explicitly.** "If a fix didn't work, say so and explain why." Agents default to reporting success otherwise.
- **Be explicit about output format.** Vague instructions get vague returns. Specify the exact report structure you expect back.
- **Require layout intent before editing.** For any layout change (repositioning, resizing, splitting elements), the agent must describe the intended new layout in plain language *before* touching any values. Structural layout problems can't be found through EMU trial-and-error — if the approach is wrong, catch it in words before the first render.
- **Cap iterations at 2 per approach, then file partial/failed.** If 2 render attempts don't fix the issue, the agent must stop, file `"status": "partial"`, and explain why it's not converging. It must not continue guessing. Slides that need structural redesign should be escalated to the user for a layout decision.
- **A report is required for every slide dispatched, including failures.** No slide may be silently abandoned. If the agent ran out of attempts or couldn't converge, the report must document what was tried and why it didn't work.

### Background agent workflow for parallel QA

Use this workflow when fixing multiple slides. Each agent works on its own single-slide deck to avoid conflicts, then a merge step combines the results.

There are two paths depending on how the deck was produced. Use the right one — do not mix them.

---

#### Path A: Code-generated deck (pptxgenjs generator exists)

**Setup:**

```bash
QA_DIR="outputs/<name>-qa"
mkdir -p "$QA_DIR/patches" "$QA_DIR/reports"
```

**Dispatch:** Launch one background agent per slide (or per 2–3 slides). Each agent:
1. Gets the generator file path, the slide number(s), and the issues found.
2. Produces a **single-slide PPTX** at `$QA_DIR/patches/slide-{N}.pptx`.
3. Writes a JSON report to `$QA_DIR/reports/slide-{N}.json`.

Use `run_in_background: true` on each Agent call. Launch all agents in a single message so they run concurrently.

**Agent prompt template (Path A):**

> Fix slide {N} from `{generator_file}`. Issues found:
> {issue_list}
>
> **Your workflow:**
> 1. Read the generator code for slide {N}.
> 2. **Before touching any values:** describe in plain language what layout change you intend to make and why it will fix the issue.
> 3. Create a standalone script at `{qa_dir}/patches/gen-slide-{N}.js` that generates **only slide {N}** as a single-slide PPTX at `{qa_dir}/patches/slide-{N}.pptx`. Copy the relevant slide code from the generator — do not modify the original generator file.
> 4. Fix the issues in your standalone script. For each change, note the original value and the new value (EMU, font size, position, etc.).
> 5. Run your script to produce the single-slide PPTX.
> 6. Render and verify:
>    ```
>    soffice --headless --convert-to pdf {qa_dir}/patches/slide-{N}.pptx --outdir {qa_dir}/patches
>    pdftoppm -jpeg -r 120 {qa_dir}/patches/slide-{N}.pdf {qa_dir}/patches/slide-{N}
>    rm {qa_dir}/patches/slide-{N}.pdf
>    ```
> 7. Read the rendered image. Describe what you see. Does the issue persist?
> 8. If the fix didn't work: adjust your approach **once more**, re-run, re-render, re-read. If it still isn't fixed after 2 attempts, stop — do not keep guessing.
> 9. Write your report to `{qa_dir}/reports/slide-{N}.json` (see report format below).
>
> **Rules:**
> - Do NOT modify the original generator file. Work only in `{qa_dir}/patches/`.
> - Do not report success unless you read the re-rendered image and confirmed the fix.
> - If after 2 attempts the fix didn't work, set `"status": "partial"` or `"failed"` and explain why in `remaining_issues`. Do not file `"status": "fixed"` unless verified.

**Merge (Path A):** Once all agents finish:
1. Read each report. Skip any with `"status": "failed"`.
2. For each `"fixed"` or `"partial"` report, port the changes from the agent's `gen-slide-{N}.js` back into the original generator file.
3. Re-run the full generator to produce the final deck.
4. Re-render and do a final inspection pass on the merged result.

This is the only step that modifies the original generator file — it happens sequentially after all agents are done, so there are no conflicts.

---

#### Path B: Existing PPTX (no generator — direct XML editing)

Do not write a `.js` generator for an existing PPTX. The overhead of re-implementing a slide from scratch in pptxgenjs outweighs the benefit, and the output will differ from the original in font rendering, theme inheritance, and image handling. Edit the XML directly instead.

**Setup:**

```bash
QA_DIR="<name>-qa"
mkdir -p "$QA_DIR/patches" "$QA_DIR/reports"
```

**Extract:** Before dispatching agents, unpack the PPTX and extract single-slide copies so agents don't conflict:

```bash
# Unpack the full deck once into a shared work directory
cp <file.pptx> "$QA_DIR/main-work.pptx"
cd "$QA_DIR" && mkdir main-work && cd main-work && unzip ../main-work.pptx && cd ../..

# For each slide N being patched, copy the full unpacked tree into a per-slide directory
cp -r "$QA_DIR/main-work" "$QA_DIR/patches/slide-{N}-work"
# Then delete all slides except slide N from that copy before dispatching
```

Or use python-pptx to produce a single-slide PPTX per agent:

```python
from pptx import Presentation

src = Presentation("<file.pptx>")
# ... keep only slide i, save as patches/slide-{i+1}.pptx
```

**Dispatch:** Each agent gets the extracted single-slide PPTX, the slide number, and the issues found.

**Agent prompt template (Path B):**

> Fix slide {N} from `{input_pptx}` (single-slide PPTX). Issues found:
> {issue_list}
>
> **Your workflow:**
> 1. Read the rendered input image (if provided) and the issue description.
> 2. **Before touching any XML:** describe in plain language what layout change you intend to make and why it will fix the issue. For example: "I will widen the title text box so the heading fits on one line, which will stop it from overlapping the description below."
> 3. Unpack the PPTX, locate the relevant shape(s) in `ppt/slides/slide1.xml`, and edit the XML. Note the original value and what you're changing it to (EMU, font size, etc.).
> 4. Repack and render:
>    ```
>    cd {work_dir} && zip -r ../slide-{N}-fixed.pptx . && cd ..
>    soffice --headless --convert-to pdf slide-{N}-fixed.pptx --outdir .
>    pdftoppm -jpeg -r 120 slide-{N}-fixed.pdf slide-{N}
>    rm slide-{N}-fixed.pdf
>    ```
> 5. Read the rendered image. Describe what you see. Does the issue persist?
> 6. If the fix didn't work: adjust your approach **once more** (including reconsidering the layout strategy), re-edit, re-render, re-read. If it still isn't fixed after 2 attempts, stop.
> 7. Clean up: remove the unpacked work directory (`{work_dir}`), keep only the `.pptx` and rendered image.
> 8. Write your report to `{qa_dir}/reports/slide-{N}.json` (see report format below).
>
> **Rules:**
> - Do NOT modify the original source PPTX. Work only in your assigned patch directory.
> - Do not report success unless you read the re-rendered image and confirmed the fix.
> - After 2 failed attempts, file `"status": "partial"` or `"failed"` — do not keep iterating.
> - Remove all unpacked XML directories when done. Only the final `.pptx` and rendered image should remain in `patches/`.

**Merge (Path B):** Once all agents finish:
1. Read each report. Skip any with `"status": "failed"`.
2. For each `"fixed"` or `"partial"` report, apply the XML changes to the main unpacked deck (`$QA_DIR/main-work/`), using the before/after values in the report as a guide.
3. Repack the main deck and render all affected slides for a final inspection.

---

#### Shared: report format

All agents (both paths) write this JSON report:

```json
{
  "slide": N,
  "status": "fixed" | "partial" | "failed",
  "patch_file": "{qa_dir}/patches/slide-{N}.pptx",
  "generator_patch": "{qa_dir}/patches/gen-slide-{N}.js",  // Path A only
  "intended_approach": "plain-language description of the layout change",
  "fixes": [
    {
      "issue": "description of what was wrong",
      "element": "code identifier or XML path",
      "before": "original value",
      "after": "new value",
      "verified": true | false,
      "notes": "what the re-rendered image shows"
    }
  ],
  "remaining_issues": ["anything still broken or why the fix didn't converge"]
}
```

---

#### Monitor, triage, cleanup

**Register tasks before dispatching.** Before launching any background agent, create a task for it with `TaskCreate`. Set status to `in_progress`. Include the slide number and issue summary in the task name so `TaskList` gives a readable view of what's running. Example:

```
TaskCreate: "QA slide 7 — pricing card overlap"  (status: in_progress)
TaskCreate: "QA slide 10 — phase cards overflow"  (status: in_progress)
```

This is the mechanism for stopping a stuck agent: use `TaskStop` on its task ID. Without a registered task, there is no handle to stop it.

**Monitor:** As agents complete (you'll be notified), check progress:

```bash
ls <qa_dir>/reports/        # which slides are done
ls <qa_dir>/patches/*.pptx  # which patches exist
```

Cross-reference with `TaskList` to confirm all agents have reported in. When an agent's report lands, mark its task complete (`TaskUpdate` → `completed`). If it filed `"status": "failed"` or `"partial"`, mark the task `failed` so it's visible.

**To stop a specific agent:** use `TaskStop` with its task ID. Then check whether a partial report or patch file was written before the agent was stopped — use whatever exists as the starting point for a re-dispatch.

**Triage:** After the merge, re-inspect. If slides still have issues, re-dispatch only those slides. Create new tasks for the re-dispatch. Include the previous report in the new agent's prompt so it doesn't repeat the same broken approach. If a slide's `remaining_issues` indicates a structural redesign decision is needed, escalate to the user before re-dispatching.

**Cleanup:** After the final deck passes QA, remove the working directory:

```bash
rm -rf <name>-qa/
```

## Dependencies

LibreOffice (`soffice`) and Poppler (`pdftoppm`) must be installed for rendering:

```bash
brew install --cask libreoffice && brew install poppler   # macOS
sudo apt install libreoffice poppler-utils                # Debian/Ubuntu
```

## Bakeoff

The `bakeoff/` directory contains an eval comparing generation with and without this QA prompt. See `bakeoff/PROMPT.md` for details.
