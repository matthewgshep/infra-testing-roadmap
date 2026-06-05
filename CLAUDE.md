# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-purpose tool that turns an Excel roadmap of Adobe Acrobat (Reader / Reduced Mode) A/B experiments into a self-contained interactive HTML dashboard published to GitHub Pages. There is no build system, test suite, or framework — it's three standalone Python scripts plus the HTML they emit.

## Two locations, one data flow

This is the critical mental model. There are **two separate git checkouts**:

- **This folder** (`~/Library/CloudStorage/OneDrive-Adobe/Reader & Reduced Mode - Infra Testing/`) — the OneDrive-synced **source**: the Excel file, the Python scripts, and generated HTML copies. Remote `origin` is `matthewgshep/infra-testing-roadmap`.
- **`~/infra-testing-roadmap/`** — a *separate clone* used purely as the **GitHub Pages deploy target**. `--deploy` copies the generated HTML into it as `index.html`, commits, and pushes.

Data flow:
```
Reader & Reduced Mode Testing Roadmap.xlsx   (source of truth)
        │  generate_gantt.py reads sheets: Results, Backlog, Planning
        ▼
Product Testing Roadmap.html                 (self-contained output, default --output)
        │  --deploy: shutil.copy2 → ~/infra-testing-roadmap/index.html → git push
        ▼
https://matthewgshep.github.io/infra-testing-roadmap/
```

The Excel workbook is the source of truth. Everything else is generated or derived from it.

## Commands

Use the project venv: `.venv/` already has `openpyxl` (and the scripts also expect `anthropic`, `watchdog`).

```bash
# Generate HTML only (reads Excel, writes "Product Testing Roadmap.html")
python generate_gantt.py

# Generate + deploy to GitHub Pages (copies to ~/infra-testing-roadmap/index.html, commits, pushes)
python generate_gantt.py --deploy

# Generate + run a local server with live Planning-Roadmap auto-save back to Excel
python generate_gantt.py --serve            # opens http://localhost:8060

# Import experiment screenshots → append rows to the Results sheet (server-side, via Claude vision)
export ANTHROPIC_API_KEY=sk-ant-...
python import_results.py --folder ~/Desktop/screenshots
python import_results.py --dry-run          # extract + print, don't write Excel

# Auto-deploy on every Excel save (filewatcher)
python watch_and_deploy.py                  # foreground; writes ~/.gantt-watcher.pid
kill $(cat ~/.gantt-watcher.pid)            # stop it
```

There are no tests or linters. Verify changes by regenerating the HTML and opening it.

## generate_gantt.py — the core (~2600 lines, one file)

`extract_data()` reads the workbook; `generate_html()` is one giant f-string that emits a fully self-contained HTML file (inline CSS + JS, data injected as a `RAW_DATA` JSON blob). Because it's an f-string, **all literal `{` and `}` in the embedded CSS/JS must be doubled (`{{` `}}`)** — this is the single easiest thing to break when editing the HTML/JS.

The emitted page has three tabs:
1. **Results Roadmap** — the Gantt of completed/running/scheduled tests.
2. **Planning Roadmap** — a drag-and-drop planner. Persists to `plan.json` in the GitHub repo via the **GitHub Contents API called directly from the browser**, using a PAT stored in `localStorage` (`gh_plan_token`). When run with `--serve`, planning instead POSTs to `/save-plan`, which writes a `Planning` sheet back into the Excel file (`save_plan_to_excel`).
3. **Import Results** — drag screenshots in; calls the **Anthropic API directly from the browser** (`anthropic-dangerous-direct-browser-access`), key in `localStorage` (`anthropic_import_key`), to extract result rows for pasting into Excel.

### Excel `Results` sheet column layout (1-indexed, header row 1)
Both `generate_gantt.py` and `import_results.py` hard-code these positions — keep them in sync:

| Col | Field | Col | Field |
|-----|-------|-----|-------|
| 2 | Product | 9 | Units % lift |
| 3 | Test Name | 10 | Winner |
| 4 | Start date | 11 | GDS link (text + hyperlink) |
| 5 | End date | 12 | Action-if-contrary |
| 6 | QGNARR | 13 | Details / decision |
| 7 | GNARR lift | 14 | Result-contrary |
| 8 | CTR lift | 15 | PM commentary |

Status is derived, not stored: no end date + start in future → **Scheduled**; no end date + started → **Running**; else **Complete**. A row is skipped unless it has a Product and a real `datetime` start. Product strings `Acrobat Reader`/`Acrobat Reduced Mode` are normalized to `Reader`/`Reduced Mode`.

### Fiscal calendar
`FY_STARTS` maps Adobe fiscal years to their start date (Saturday nearest Dec 1). `_extend_fy_starts()` auto-extends 3 FYs past today using a flat 364-day (52-week) roll. Adobe occasionally has 53-week fiscal years, so these auto-generated future starts drift and need manual correction roughly every ~5 years.

## Gotchas

- **Output filename mismatch by design**: default `--output` is `Product Testing Roadmap.html`, but `--deploy` copies it to the repo as `index.html`. The folder therefore accumulates several near-identical HTML copies (`index.html`, `Product_Testing_Roadmap.html`, `Product Testing Roadmap.html`) — don't assume they're meaningfully different.
- `plan.json` lives in both checkouts and is committed alongside `index.html` on deploy; it's the planning persistence file, not config.
- `generate_gantt.py` is currently untracked in this repo's git.
- `~$...xlsx` is an Excel lock file — ignore it. Editing the workbook in Excel while a script reads it can cause inconsistent reads.
- Embedded model IDs are pinned in source: `import_results.py` uses `claude-opus-4-5`; the browser Import tab uses `claude-sonnet-4-20250514`. Update both if migrating models.
