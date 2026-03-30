# Bakeoff Orchestration

## What this does

Compares three Claude Code PPTX skills by generating decks from identical prompts and comparing the outputs side-by-side.

| Skill | Source | License |
|-------|--------|---------|
| deck-builder | This repo | MIT |
| MiniMax pptx-generator | [MiniMax-AI/skills](https://github.com/MiniMax-AI/skills/tree/main/skills/pptx-generator) | MIT |
| Anthropic pptx | [anthropics/skills](https://github.com/anthropics/skills/tree/main/skills/pptx) | Proprietary |

## How it works

`run-bakeoff.sh` uses `claude --print` to invoke Claude Code in headless mode, injecting each skill's docs as a system prompt. Each invocation gets a clean working directory with pptxgenjs pre-installed.

For each prompt, all 3 skills run **in parallel**. Each skill writes a Node.js generator, runs it, and produces a `.pptx` file.

## Quick start

```bash
cd bakeoff/

# 1. Clone skill repos + install pptxgenjs
./run-bakeoff.sh setup

# 2. Generate decks (all prompts, all skills)
./run-bakeoff.sh generate

# 3. Render to slide images
./run-bakeoff.sh render

# 4. See what was produced
./run-bakeoff.sh summary
```

### Run a single prompt

```bash
./run-bakeoff.sh generate 01       # just 01-creative
./run-bakeoff.sh generate strategy  # just 03-strategy
```

### Re-run from scratch

```bash
./run-bakeoff.sh clean
./run-bakeoff.sh generate
```

## Directory layout

After running:

```
bakeoff/
├── prompts/              # Shared input (checked in)
│   ├── 01-creative.md
│   ├── 02-software.md
│   └── 03-strategy.md
├── deck-builder/         # Our skill's outputs (gitignored)
│   ├── 01-creative/
│   │   ├── gen-deck.js
│   │   ├── output/*.pptx
│   │   └── output/slide-*.jpg
│   └── ...
├── minimax/              # MiniMax outputs (gitignored)
├── anthropic/            # Anthropic outputs (gitignored)
├── .vendor/              # Cloned skill repos (gitignored)
├── SCORECARD.md          # Rating template
├── ORCHESTRATION.md      # This file
└── run-bakeoff.sh        # Automation script
```

## Scoring

After rendering, fill in `SCORECARD.md` with ratings across 5 dimensions per prompt. Open slide images side-by-side for visual comparison.

## Environment variables

| Variable | Default | Description |
|----------|---------|-------------|
| `BAKEOFF_MAX_BUDGET` | `5` | Max USD per Claude Code invocation |
