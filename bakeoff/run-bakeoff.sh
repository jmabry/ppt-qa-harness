#!/bin/bash
set -euo pipefail

# Bakeoff: Compare deck-builder vs MiniMax vs Anthropic pptx skills
#
# Generates PPTX decks from identical prompts using 3 different Claude Code
# skills, then renders outputs to images for side-by-side comparison.
#
# Prerequisites:
#   - claude CLI (Claude Code 2.x+)
#   - npm (for pptxgenjs install)
#   - soffice (LibreOffice) + pdftoppm (Poppler) for rendering
#   - git (for cloning skill repos)
#
# Usage:
#   ./run-bakeoff.sh setup          # clone skills, install deps
#   ./run-bakeoff.sh generate [N]   # generate decks (optionally just prompt N)
#   ./run-bakeoff.sh render         # convert all PPTX to slide images
#   ./run-bakeoff.sh summary        # print comparison summary
#   ./run-bakeoff.sh clean          # remove generated outputs (keep prompts)

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
PROMPTS_DIR="$SCRIPT_DIR/prompts"
VENDOR_DIR="$SCRIPT_DIR/.vendor"

# Skill sources — cloned into .vendor/ during setup
DECKBUILDER_SKILL="$REPO_ROOT"
MINIMAX_SKILL="$VENDOR_DIR/minimax-skills/skills/pptx-generator"
ANTHROPIC_SKILL="$VENDOR_DIR/anthropic-skills/skills/pptx"

SKILLS=(deck-builder minimax anthropic)
PROMPTS=(01-creative 02-software 03-strategy)
MAX_BUDGET="${BAKEOFF_MAX_BUDGET:-5}"  # USD per invocation

# ---- Helpers ----

log() { echo "[$(date +%H:%M:%S)] $*"; }

load_skill_docs() {
  local skill="$1"
  case "$skill" in
    deck-builder)
      cat "$DECKBUILDER_SKILL/SKILL.md" \
          "$DECKBUILDER_SKILL/architecture.md" \
          "$DECKBUILDER_SKILL/pptxgenjs.md"
      ;;
    minimax)
      if [ ! -f "$MINIMAX_SKILL/SKILL.md" ]; then
        echo "ERROR: MiniMax skill not found — run '$0 setup' first" >&2
        return 1
      fi
      cat "$MINIMAX_SKILL/SKILL.md" "$MINIMAX_SKILL"/references/*.md
      ;;
    anthropic)
      if [ ! -f "$ANTHROPIC_SKILL/SKILL.md" ]; then
        echo "ERROR: Anthropic skill not found — run '$0 setup' first" >&2
        return 1
      fi
      cat "$ANTHROPIC_SKILL/SKILL.md" "$ANTHROPIC_SKILL/pptxgenjs.md"
      ;;
  esac
}

# ---- Commands ----

cmd_setup() {
  log "Setting up bakeoff environment..."
  mkdir -p "$VENDOR_DIR"

  # Clone MiniMax skill (MIT)
  if [ ! -d "$VENDOR_DIR/minimax-skills" ]; then
    log "Cloning MiniMax skills from GitHub (MIT licensed)..."
    git clone --depth 1 https://github.com/MiniMax-AI/skills.git "$VENDOR_DIR/minimax-skills"
  else
    log "MiniMax skills already cloned — pulling latest..."
    git -C "$VENDOR_DIR/minimax-skills" pull --ff-only 2>/dev/null || true
  fi

  # Clone Anthropic skill (proprietary — source-available for reference)
  if [ ! -d "$VENDOR_DIR/anthropic-skills" ]; then
    log "Cloning Anthropic skills from GitHub (proprietary — for comparison only)..."
    git clone --depth 1 https://github.com/anthropics/skills.git "$VENDOR_DIR/anthropic-skills"
  else
    log "Anthropic skills already cloned — pulling latest..."
    git -C "$VENDOR_DIR/anthropic-skills" pull --ff-only 2>/dev/null || true
  fi

  # Verify skills exist
  for skill in deck-builder minimax anthropic; do
    if load_skill_docs "$skill" > /dev/null 2>&1; then
      log "OK: $skill skill found"
    else
      log "WARNING: $skill skill not found"
    fi
  done

  # Create output directories and install pptxgenjs
  for skill in "${SKILLS[@]}"; do
    for prompt in "${PROMPTS[@]}"; do
      local dir="$SCRIPT_DIR/$skill/$prompt"
      mkdir -p "$dir/output"
      if [ ! -d "$dir/node_modules/pptxgenjs" ]; then
        log "Installing pptxgenjs in $skill/$prompt..."
        npm install --prefix "$dir" pptxgenjs 2>/dev/null
      fi
    done
  done

  log "Setup complete."
}

run_one() {
  local skill="$1"
  local prompt_name="$2"
  local work_dir="$SCRIPT_DIR/$skill/$prompt_name"
  local prompt_file="$PROMPTS_DIR/$prompt_name.md"
  local log_file="$work_dir/generation.log"

  if [ ! -f "$prompt_file" ]; then
    log "ERROR: Prompt file not found: $prompt_file"
    return 1
  fi

  # Skip if PPTX already exists
  if ls "$work_dir"/output/*.pptx 1>/dev/null 2>&1; then
    log "SKIP $skill/$prompt_name — PPTX already exists (use 'clean' to reset)"
    return 0
  fi

  local skill_docs
  skill_docs=$(load_skill_docs "$skill") || return 1

  local prompt_text
  prompt_text=$(cat "$prompt_file")

  log "START $skill/$prompt_name"

  claude --print \
    --system-prompt "$skill_docs" \
    --allowedTools "Bash Write Read Glob Grep" \
    --max-budget-usd "$MAX_BUDGET" \
    -d "$work_dir" \
    "${prompt_text}

Generate the deck as a Node.js script using pptxgenjs (already installed in node_modules/).
Write the generator to gen-deck.js, run it, and save the .pptx to the output/ directory.
Do not install any packages — pptxgenjs is pre-installed. Use require('pptxgenjs') — it resolves from node_modules/." \
    > "$log_file" 2>&1

  local exit_code=$?
  if [ $exit_code -eq 0 ] && ls "$work_dir"/output/*.pptx 1>/dev/null 2>&1; then
    log "DONE $skill/$prompt_name — $(ls "$work_dir"/output/*.pptx)"
  else
    log "FAIL $skill/$prompt_name (exit=$exit_code) — see $log_file"
  fi
}

cmd_generate() {
  local filter="${1:-}"

  # Verify setup
  if [ ! -d "$VENDOR_DIR" ]; then
    log "ERROR: Run '$0 setup' first"
    exit 1
  fi

  for prompt_name in "${PROMPTS[@]}"; do
    # Filter to specific prompt if requested
    if [ -n "$filter" ] && [[ "$prompt_name" != *"$filter"* ]]; then
      continue
    fi

    log "=== Prompt: $prompt_name ==="

    # Launch all 3 skills in parallel for this prompt
    local pids=()
    for skill in "${SKILLS[@]}"; do
      run_one "$skill" "$prompt_name" &
      pids+=($!)
    done

    # Wait for all 3 to finish
    local any_failed=0
    for pid in "${pids[@]}"; do
      wait "$pid" || any_failed=1
    done

    if [ "$any_failed" -eq 1 ]; then
      log "WARNING: Some skills failed for $prompt_name — check logs"
    fi
  done

  log "Generation complete."
}

cmd_render() {
  log "Rendering all PPTX files to images..."

  for skill in "${SKILLS[@]}"; do
    for prompt_name in "${PROMPTS[@]}"; do
      local dir="$SCRIPT_DIR/$skill/$prompt_name/output"
      for pptx in "$dir"/*.pptx; do
        [ -f "$pptx" ] || continue
        local base
        base=$(basename "$pptx" .pptx)

        # Skip if already rendered
        if ls "$dir"/slide-*.jpg 1>/dev/null 2>&1; then
          log "SKIP render $skill/$prompt_name — images exist"
          continue
        fi

        log "Rendering $skill/$prompt_name..."
        soffice --headless --convert-to pdf "$pptx" --outdir "$dir" 2>/dev/null || {
          log "FAIL render $skill/$prompt_name"
          continue
        }

        local pdf_file="$dir/$base.pdf"
        if [ -f "$pdf_file" ]; then
          pdftoppm -jpeg -r 150 "$pdf_file" "$dir/slide"
          local count
          count=$(ls "$dir"/slide-*.jpg 2>/dev/null | wc -l | tr -d ' ')
          log "DONE render $skill/$prompt_name — $count slides"
        fi
      done
    done
  done
}

cmd_summary() {
  echo ""
  echo "============================================"
  echo "  BAKEOFF RESULTS"
  echo "============================================"
  echo ""

  for prompt_name in "${PROMPTS[@]}"; do
    echo "### $prompt_name"
    echo ""
    printf "  %-18s %8s %8s\n" "Skill" "PPTX" "Slides"
    printf "  %-18s %8s %8s\n" "-----" "----" "------"

    for skill in "${SKILLS[@]}"; do
      local dir="$SCRIPT_DIR/$skill/$prompt_name"
      local pptx_count=0
      local slide_count=0
      [ -d "$dir" ] && pptx_count=$(find "$dir" -name "*.pptx" 2>/dev/null | wc -l | tr -d ' ')
      [ -d "$dir" ] && slide_count=$(find "$dir" -name "slide-*.jpg" 2>/dev/null | wc -l | tr -d ' ')
      printf "  %-18s %8s %8s\n" "$skill" "$pptx_count" "$slide_count"
    done
    echo ""
  done

  echo "To compare visually:"
  echo "  open $SCRIPT_DIR/deck-builder/*/output/slide-*.jpg"
  echo "  open $SCRIPT_DIR/minimax/*/output/slide-*.jpg"
  echo "  open $SCRIPT_DIR/anthropic/*/output/slide-*.jpg"
  echo ""
  echo "Score the results in: $SCRIPT_DIR/SCORECARD.md"
}

cmd_clean() {
  log "Cleaning generated outputs..."
  for skill in "${SKILLS[@]}"; do
    for prompt_name in "${PROMPTS[@]}"; do
      local dir="$SCRIPT_DIR/$skill/$prompt_name"
      rm -f "$dir"/output/*.pptx "$dir"/output/*.pdf "$dir"/output/slide-*.jpg
      rm -f "$dir"/gen-*.js "$dir"/generation.log
      log "Cleaned $skill/$prompt_name"
    done
  done
  log "Done. Run 'generate' to re-create."
}

# ---- Main ----

case "${1:-help}" in
  setup)    cmd_setup ;;
  generate) cmd_generate "${2:-}" ;;
  render)   cmd_render ;;
  summary)  cmd_summary ;;
  clean)    cmd_clean ;;
  help|*)
    echo "Usage: $0 <command> [args]"
    echo ""
    echo "Commands:"
    echo "  setup              Clone skill repos from GitHub, install pptxgenjs"
    echo "  generate [filter]  Generate decks (filter: '01', 'creative', etc.)"
    echo "  render             Convert all PPTX to slide images via LibreOffice"
    echo "  summary            Print comparison of what was generated"
    echo "  clean              Remove generated files (keep prompts and deps)"
    echo ""
    echo "Skills compared:"
    echo "  deck-builder  — this repo's skill (MIT)"
    echo "  minimax       — MiniMax pptx-generator from github.com/MiniMax-AI/skills (MIT)"
    echo "  anthropic     — Anthropic pptx from github.com/anthropics/skills (proprietary)"
    echo ""
    echo "Environment variables:"
    echo "  BAKEOFF_MAX_BUDGET  Max USD per Claude invocation (default: 5)"
    ;;
esac
