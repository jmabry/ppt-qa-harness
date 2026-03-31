#!/bin/bash
set -euo pipefail

# Bakeoff: Compare harness-guided vs unguided pptx generation
#
# Runs the same Anthropic pptx skill on identical prompts under two conditions:
#   "generate" — skill docs only, Claude decides whether to QA its own output
#   "harness"  — skill docs + CLAUDE.md, which mandates a pptx-qa agent
#                loop until the output passes visual inspection
#
# The only variable is the harness instruction. Same model, same tools, same
# prompts. This isolates the effect of structured QA on output quality.
#
# Prerequisites:
#   claude (Claude Code CLI)    jq            node / npm
#   soffice (LibreOffice)       pdftoppm      git
#   Run `./run-bakeoff.sh setup` to verify deps and fetch skill docs.
#
# Directory layout:
#   prompts/                 One .md file per deck topic
#   generators/              Generated Node.js scripts ({topic}-{tag}.js)
#   outputs/                 PPTX, PDF, slide JPEGs, and logs
#   bakeoff.config.json      Shared defaults (committed)
#   bakeoff.config.local.json  Personal overrides (gitignored)
#
# Config (bakeoff.config.json):
#   prompts          Array of prompt names matching files in prompts/
#   max_budget_usd   Max spend per Claude invocation (default: 10)
#   model            Model override, or null for default
#   parallel         true to run prompts concurrently, false for sequential
#   claude_args.{generate,harness}:
#     allowed_tools    Space-separated tool names passed to --allowedTools
#     permission_mode  Permission mode (null = default, "bypassPermissions" etc.)
#
# To override config locally without affecting the repo, create
# bakeoff.config.local.json — it deep-merges over the base config.
# Example (skip permission prompts for your own runs):
#   { "claude_args": { "generate": { "permission_mode": "bypassPermissions" },
#                      "harness":  { "permission_mode": "bypassPermissions" } } }
#
# Commands:
#   ./run-bakeoff.sh                    # default: runs "all"
#   ./run-bakeoff.sh all [filter]       # setup + generate + harness + render + summary
#   ./run-bakeoff.sh setup              # clone skill, install deps, smoke-test claude
#   ./run-bakeoff.sh generate [filter]  # generate decks without harness
#   ./run-bakeoff.sh harness  [filter]  # generate decks with harness (QA loop)
#   ./run-bakeoff.sh render             # convert PPTX → PDF → slide JPEGs
#   ./run-bakeoff.sh summary            # print side-by-side comparison table
#   ./run-bakeoff.sh clean              # remove outputs + generators (keeps prompts)
#
# The optional [filter] is a substring match on prompt names, e.g.:
#   ./run-bakeoff.sh generate corporate   # only run the corporate prompt
#
# Scoring: after both runs complete, fill in SCORECARD.md with visual ratings.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
PROMPTS_DIR="$SCRIPT_DIR/prompts"
GENERATORS_DIR="$SCRIPT_DIR/generators"
OUTPUTS_DIR="$SCRIPT_DIR/outputs"
VENDOR_DIR="$SCRIPT_DIR/.vendor"
CLAUDE_DIR="$REPO_ROOT/.claude"
ANTHROPIC_SKILL="$VENDOR_DIR/anthropic-skills/skills/pptx"

CONFIG="$SCRIPT_DIR/bakeoff.config.json"
LOCAL_CONFIG="$SCRIPT_DIR/bakeoff.config.local.json"

# ---- Config (jq merges base + local overrides) ----

cfg() {
  local key="$1"
  if [ -f "$LOCAL_CONFIG" ]; then
    jq -rs ".[0] * .[1] | $key // empty" "$CONFIG" "$LOCAL_CONFIG"
  else
    jq -r "$key // empty" "$CONFIG"
  fi
}

cfg_array() {
  local key="$1"
  if [ -f "$LOCAL_CONFIG" ]; then
    jq -rs ".[0] * .[1] | $key // [] | .[]" "$CONFIG" "$LOCAL_CONFIG"
  else
    jq -r "$key // [] | .[]" "$CONFIG"
  fi
}

# ---- Load config ----

PROMPTS=()
while IFS= read -r line; do PROMPTS+=("$line"); done < <(cfg_array '.prompts')
MAX_BUDGET="$(cfg '.max_budget_usd')"
PARALLEL="$(cfg '.parallel')"
MODEL="$(cfg '.model')"
CLAUDE_CMD="$(cfg '.claude_command')"
CLAUDE_CMD="${CLAUDE_CMD:-claude}"  # default to 'claude' if not set

log() { echo "[$(date +%H:%M:%S)] $*"; }

load_anthropic_docs() {
  if [ ! -f "$ANTHROPIC_SKILL/SKILL.md" ]; then
    echo "ERROR: Anthropic skill not found — run '$0 setup' first" >&2; return 1
  fi
  cat "$ANTHROPIC_SKILL/SKILL.md" "$ANTHROPIC_SKILL/pptxgenjs.md"
}

# Build the claude CLI invocation array for a given mode
build_claude_args() {
  local mode="$1"  # "generate" or "harness"
  local args=(--print)

  local allowed_tools
  allowed_tools="$(cfg ".claude_args.${mode}.allowed_tools")"
  [ -n "$allowed_tools" ] && args+=(--allowedTools "$allowed_tools")

  local permission_mode
  permission_mode="$(cfg ".claude_args.${mode}.permission_mode")"
  [ -n "$permission_mode" ] && args+=(--permission-mode "$permission_mode")

  [ -n "$MAX_BUDGET" ] && args+=(--max-budget-usd "$MAX_BUDGET")
  [ -n "$MODEL" ]      && args+=(--model "$MODEL")

  # Return via global — bash can't return arrays
  CLAUDE_ARGS=("${args[@]}")
}

# ---- Commands ----

cmd_setup() {
  log "Setting up bakeoff environment..."
  mkdir -p "$VENDOR_DIR" "$GENERATORS_DIR" "$OUTPUTS_DIR"

  # Check dependencies
  local missing=()
  command -v jq       >/dev/null || missing+=(jq)
  command -v claude   >/dev/null || missing+=(claude)
  command -v npm      >/dev/null || missing+=(npm)
  command -v node     >/dev/null || missing+=(node)
  command -v soffice  >/dev/null || missing+=(soffice)
  command -v pdftoppm >/dev/null || missing+=(pdftoppm)
  command -v git      >/dev/null || missing+=(git)
  if [ "${#missing[@]}" -gt 0 ]; then
    log "ERROR: Missing dependencies: ${missing[*]}"
    log "  brew install ${missing[*]}"
    exit 1
  fi
  log "OK: all dependencies found"

  if [ ! -d "$VENDOR_DIR/anthropic-skills" ]; then
    log "Cloning Anthropic skills from GitHub..."
    git clone --depth 1 https://github.com/anthropics/skills.git "$VENDOR_DIR/anthropic-skills"
  else
    log "Anthropic skills already cloned — pulling latest..."
    git -C "$VENDOR_DIR/anthropic-skills" pull --ff-only 2>/dev/null || true
  fi

  if load_anthropic_docs > /dev/null 2>&1; then
    log "OK: anthropic skill found"
  else
    log "WARNING: anthropic skill not found"
  fi

  if [ ! -d "$SCRIPT_DIR/node_modules/pptxgenjs" ]; then
    log "Installing pptxgenjs..."
    npm install --prefix "$SCRIPT_DIR" pptxgenjs 2>/dev/null || {
      log "WARNING: npm install failed — check permissions or run manually"
    }
  fi

  log "Setup complete."
}

run_one() {
  local mode="$1" prompt_name="$2"
  local tag="$( [ "$mode" = "harness" ] && echo "qa" || echo "baseline" )"
  local output_file="$OUTPUTS_DIR/${prompt_name}-${tag}.pptx"
  local log_file="$OUTPUTS_DIR/${prompt_name}-${tag}.log"
  local meta_file="$OUTPUTS_DIR/${prompt_name}-${tag}.meta.json"
  local prompt_file="$PROMPTS_DIR/${prompt_name}.md"

  if [ ! -f "$prompt_file" ]; then
    log "ERROR: Prompt file not found: $prompt_file"; return 1
  fi
  local gen_file="$GENERATORS_DIR/${prompt_name}-${tag}.js"

  if [ -f "$output_file" ]; then
    log "SKIP ${prompt_name}-${tag} — PPTX exists (use 'clean' to reset)"; return 0
  fi

  # Clean slate — remove stale generator so Claude builds fresh
  rm -f "$gen_file"

  local skill_docs prompt_text system_prompt
  skill_docs=$(load_anthropic_docs) || return 1
  prompt_text=$(cat "$prompt_file")
  system_prompt="$skill_docs"

  # Harness mode: append harness instructions + install agent
  if [ "$mode" = "harness" ]; then
    local harness_instructions
    harness_instructions=$(cat "$REPO_ROOT/CLAUDE.md")
    system_prompt="${system_prompt}

---

${harness_instructions}"
  fi

  build_claude_args "$mode"
  local start_ts
  start_ts=$(date +%s)
  log "START ${prompt_name}-${tag}"

  local exit_code=0
  (cd "$SCRIPT_DIR" && $CLAUDE_CMD "${CLAUDE_ARGS[@]}" \
    --system-prompt "$system_prompt" \
    "${prompt_text}

Generate the deck as a Node.js script using pptxgenjs (already installed in node_modules/).
Write the generator to generators/${prompt_name}-${tag}.js, run it, and save the .pptx to outputs/${prompt_name}-${tag}.pptx.
Do not install any packages — pptxgenjs is pre-installed. Use require('pptxgenjs') — it resolves from node_modules/.
If soffice or pdftoppm are not found, install them (apt install -y libreoffice poppler-utils). Use them to render and visually verify the output.
Each time you regenerate the .pptx after QA fixes, copy the previous version to outputs/${prompt_name}-${tag}-pass{N}.pptx (pass1, pass2, ...) before overwriting, so every QA iteration is preserved.
When finished, print a single JSON summary line: {\"qa_passes\": N, \"issues_found\": N, \"issues_fixed\": N, \"status\": \"CLEAN|REMAINING\"}") \
    2>&1 | grep -v "^Checking for Claude updates" \
    | grep -v "^Container already running" \
    | grep -v "^Launching Claude" \
    | grep -v "^Starting container" \
    | grep -v "Sandbox disabled" \
    | grep -v "Commands will run WITHOUT" \
    | grep -v "^$" \
    > "$log_file" || exit_code=$?

  local end_ts duration_s
  end_ts=$(date +%s)
  duration_s=$((end_ts - start_ts))
  local qa_passes
  qa_passes=$(ls "$OUTPUTS_DIR"/${prompt_name}-${tag}-pass*.pptx 2>/dev/null | wc -l | tr -d ' ')

  # Write structured metadata
  cat > "$meta_file" <<METAEOF
{
  "prompt": "${prompt_name}",
  "mode": "${mode}",
  "tag": "${tag}",
  "duration_s": ${duration_s},
  "qa_passes": ${qa_passes},
  "exit_code": ${exit_code},
  "pptx_produced": $([ -f "$output_file" ] && echo "true" || echo "false"),
  "timestamp": "$(date -u +%Y-%m-%dT%H:%M:%SZ)"
}
METAEOF

  if [ $exit_code -eq 0 ] && [ -f "$output_file" ]; then
    log "DONE ${prompt_name}-${tag} — ${duration_s}s, ${qa_passes} QA passes"
  else
    log "FAIL ${prompt_name}-${tag} — ${duration_s}s (exit=$exit_code) — see $log_file"
  fi
}

cmd_run() {
  local mode="$1" filter="${2:-}"
  if [ ! -d "$VENDOR_DIR" ]; then log "ERROR: Run '$0 setup' first"; exit 1; fi
  if [ "$mode" = "harness" ] && [ ! -f "$CLAUDE_DIR/agents/pptx-qa.md" ]; then
    log "ERROR: .claude/agents/pptx-qa.md not found"; exit 1
  fi

  # Install agent once before parallel runs (not per-run)
  if [ "$mode" = "harness" ]; then
    local agents_dir="$SCRIPT_DIR/.claude/agents"
    mkdir -p "$agents_dir"
    cp "$CLAUDE_DIR/agents/pptx-qa.md" "$agents_dir/pptx-qa.md"
  fi

  local pids=()
  for prompt_name in "${PROMPTS[@]}"; do
    if [ -n "$filter" ] && [[ "$prompt_name" != *"$filter"* ]]; then continue; fi
    if [ "$PARALLEL" = "true" ]; then
      run_one "$mode" "$prompt_name" &
      pids+=($!)
    else
      run_one "$mode" "$prompt_name"
    fi
  done

  if [ "${#pids[@]}" -gt 0 ]; then
    local any_failed=0
    for pid in "${pids[@]}"; do wait "$pid" || any_failed=1; done
    [ "$any_failed" -eq 1 ] && log "WARNING: some prompts failed — check logs"
  fi

  # Clean up agent after all runs complete
  if [ "$mode" = "harness" ]; then
    rm -f "$SCRIPT_DIR/.claude/agents/pptx-qa.md"
    rmdir "$SCRIPT_DIR/.claude/agents" 2>/dev/null || true
    rmdir "$SCRIPT_DIR/.claude" 2>/dev/null || true
  fi

  log "${mode} run complete."
}

cmd_render() {
  log "Rendering all PPTX files to images..."
  for pptx in "$OUTPUTS_DIR"/*.pptx; do
    [ -f "$pptx" ] || continue
    local base
    base=$(basename "$pptx" .pptx)
    if ls "$OUTPUTS_DIR"/${base}-slide-*.jpg 1>/dev/null 2>&1; then
      log "SKIP render $base — images exist"; continue
    fi
    log "Rendering $base..."
    soffice --headless --convert-to pdf "$pptx" --outdir "$OUTPUTS_DIR" 2>/dev/null || {
      log "FAIL render $base"; continue
    }
    local pdf_file="$OUTPUTS_DIR/$base.pdf"
    if [ -f "$pdf_file" ]; then
      pdftoppm -jpeg -r 150 "$pdf_file" "$OUTPUTS_DIR/${base}-slide" 2>/dev/null || {
        log "FAIL pdftoppm $base"; continue
      }
      local count
      count=$(ls "$OUTPUTS_DIR"/${base}-slide-*.jpg 2>/dev/null | wc -l | tr -d ' ')
      log "DONE render $base — $count slides"
    fi
  done
}

cmd_summary() {
  echo ""
  echo "============================================"
  echo "  BAKEOFF RESULTS — baseline vs QA harness"
  echo "============================================"
  echo ""

  printf "  %-15s %8s %8s %8s %8s %10s\n" "Prompt" "base" "base(s)" "qa" "qa(s)" "qa passes"
  printf "  %-15s %8s %8s %8s %8s %10s\n" "-------" "----" "-------" "---" "-----" "---------"

  for prompt_name in "${PROMPTS[@]}"; do
    local b_ok="—" b_dur="—" q_ok="—" q_dur="—" q_passes="—"

    if [ -f "$OUTPUTS_DIR/${prompt_name}-baseline.meta.json" ]; then
      b_ok=$(jq -r 'if .pptx_produced then "yes" else "FAIL" end' "$OUTPUTS_DIR/${prompt_name}-baseline.meta.json")
      b_dur="$(jq -r '.duration_s' "$OUTPUTS_DIR/${prompt_name}-baseline.meta.json")s"
    fi
    if [ -f "$OUTPUTS_DIR/${prompt_name}-qa.meta.json" ]; then
      q_ok=$(jq -r 'if .pptx_produced then "yes" else "FAIL" end' "$OUTPUTS_DIR/${prompt_name}-qa.meta.json")
      q_dur="$(jq -r '.duration_s' "$OUTPUTS_DIR/${prompt_name}-qa.meta.json")s"
      q_passes=$(jq -r '.qa_passes' "$OUTPUTS_DIR/${prompt_name}-qa.meta.json")
    fi

    printf "  %-15s %8s %8s %8s %8s %10s\n" "$prompt_name" "$b_ok" "$b_dur" "$q_ok" "$q_dur" "$q_passes"
  done

  echo ""
  echo "Generators:"
  echo "  Baseline: $(ls $GENERATORS_DIR/*-baseline.js 2>/dev/null | wc -l | tr -d ' ') files"
  echo "  QA:       $(ls $GENERATORS_DIR/*-qa.js 2>/dev/null | wc -l | tr -d ' ') files"
  echo ""
  echo "Score the results in: $SCRIPT_DIR/SCORECARD.md"
}

cmd_all() {
  local filter="${1:-}"
  cmd_setup
  log "=== Phase 1: generate (without harness) ==="
  cmd_run generate "$filter"
  log "=== Phase 2: harness (with harness) ==="
  cmd_run harness "$filter"
  log "=== Phase 3: render ==="
  cmd_render
  cmd_summary
}

cmd_clean() {
  log "Cleaning outputs and generators..."
  rm -f "$OUTPUTS_DIR"/*.pptx "$OUTPUTS_DIR"/*.pdf "$OUTPUTS_DIR"/*.jpg "$OUTPUTS_DIR"/*.log "$OUTPUTS_DIR"/*.meta.json
  rm -rf "$OUTPUTS_DIR"/*-slides
  rm -f "$GENERATORS_DIR"/*.js
  log "Done."
}

# ---- Main ----

case "${1:-all}" in
  all)      cmd_all "${2:-}" ;;
  setup)    cmd_setup ;;
  generate) cmd_run generate "${2:-}" ;;
  harness)  cmd_run harness "${2:-}" ;;
  render)   cmd_render ;;
  summary)  cmd_summary ;;
  clean)    cmd_clean "${2:-}" ;;
  help|*)
    echo "Usage: $0 <command> [args]"
    echo ""
    echo "Commands:"
    echo "  all      [filter]   Run everything: setup + generate + harness + render + summary"
    echo "  setup               Clone Anthropic skill, install pptxgenjs"
    echo "  generate [filter]   Generate decks without harness"
    echo "  harness  [filter]   Generate decks with harness (QA loop)"
    echo "  render              Convert PPTX to slide images via LibreOffice"
    echo "  summary             Print side-by-side comparison"
    echo "  clean               Remove generated files (keep prompts and deps)"
    echo ""
    echo "Config: bakeoff.config.json        (shared defaults)"
    echo "        bakeoff.config.local.json  (personal overrides, gitignored)"
    ;;
esac
