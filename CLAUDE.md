# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-purpose tool that turns an Excel roadmap of Adobe Acrobat (Reader / Reduced Mode) A/B experiments into a self-contained interactive HTML dashboard published to GitHub Pages. There is no build system, test suite, or framework — it's three standalone Python scripts plus the HTML they emit.

## Two locations, one data flow

This is the critical mental model. There are **two checkouts of the same repo (`matthewgshep/infra-testing-roadmap`), on two different branches**:

- **This folder** (`~/Library/CloudStorage/OneDrive-Adobe/Reader & Reduced Mode - Infra Testing/`) — the OneDrive-synced **source**, on branch **`main`**: the Excel file, the Python scripts, CLAUDE.md. Locally generated HTML is gitignored, not committed.
- **`~/infra-testing-roadmap/`** — the **publish clone**, kept on branch **`gh-pages`**. `--deploy` copies the generated HTML into it as `index.html`, commits, and pushes to `gh-pages`. GitHub Pages serves from `gh-pages`.

The branch split is deliberate: `main` (source) and `gh-pages` (published `index.html`) never touch the same files, so the two checkouts never collide. Don't make `main` track `index.html`/`plan.json` again, and don't point the deploy at `main`.

Data flow:
```
Reader & Reduced Mode Testing Roadmap.xlsx   (source of truth, on main)
        │  generate_gantt.py reads sheets: Results, Backlog, Planning
        ▼
Product Testing Roadmap.html                 (self-contained output, default --output; gitignored)
        │  --deploy: shutil.copy2 → ~/infra-testing-roadmap/index.html → commit + push to gh-pages
        ▼
https://matthewgshep.github.io/infra-testing-roadmap/   (served from gh-pages)
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

# Import experiment screenshots → append rows to the Results sheet (Claude vision + xlsx enrichment)
export ANTHROPIC_API_KEY=sk-ant-...
python import_results.py                     # default: the results/ folder
python import_results.py --dry-run          # extract + enrich + print, don't write Excel

# Seed running experiments from Jira → Results sheet (product, test name, start date only)
export JIRA_BASE_URL=...                      # + JIRA_PAT (Server/DC) or JIRA_EMAIL/JIRA_API_TOKEN (Cloud)
python jira_sync.py --list-fields            # discover custom field ids (e.g. the start-date field)
python jira_sync.py --dry-run                # show what would be added, write nothing
python jira_sync.py                          # append new running experiments (dedup on RGS id + product)

# Auto-deploy on every Excel save (filewatcher)
python watch_and_deploy.py                  # foreground; writes ~/.gantt-watcher.pid
kill $(cat ~/.gantt-watcher.pid)            # stop it
```

There are no tests or linters. Verify changes by regenerating the HTML and opening it.

## generate_gantt.py — the core (~2600 lines, one file)

`extract_data()` reads the workbook; `generate_html()` is one giant f-string that emits a fully self-contained HTML file (inline CSS + JS, data injected as a `RAW_DATA` JSON blob). Because it's an f-string, **all literal `{` and `}` in the embedded CSS/JS must be doubled (`{{` `}}`)** — this is the single easiest thing to break when editing the HTML/JS.

The emitted page's main tabs:
- **Results Roadmap** — the Gantt of completed/running/scheduled tests.
- **Experiment Lift** / **Experiment List** / **Velocity** — analysis views over the same `Results` data.
- **Planning Roadmap** — a drag-and-drop planner. Persists to `plan.json` **on the `gh-pages` branch** via the **GitHub Contents API called directly from the browser**, using a PAT stored in `localStorage` (`gh_plan_token`). The read (`?ref=gh-pages`), the write (`branch: gh-pages` in the PUT body), and the deploy all target `gh-pages` — keep them in sync (`GH_BRANCH` constant in the JS, `PUBLISH_BRANCH` in Python). When run with `--serve`, planning instead POSTs to `/save-plan`, which writes a `Planning` sheet back into the Excel file (`save_plan_to_excel`).

Screenshot → Results import is handled **locally** (see "Ingesting result screenshots" below) — either in-session by Claude reading the images directly, or via `import_results.py`. There used to be an in-browser "Import Results" tab that called the Anthropic API directly from the page; it was removed — do not reintroduce browser-side API calls.

### Excel `Results` sheet column layout (1-indexed, header row 1)
Both `generate_gantt.py` and `import_results.py` hard-code these positions — keep them in sync:

| Col | Field | Col | Field |
|-----|-------|-----|-------|
| 2 | Product | 9 | Units % lift |
| 3 | Test Name | 10 | Winner |
| 4 | Start date | 11 | GDS link (text + hyperlink) |
| 5 | End date | 12 | Decision/Details |
| 6 | QGNARR | 13 | Result contrary to expectation |
| 7 | GNARR lift | 14 | PM Questions/Commentary |
| 8 | CTR lift | | *(sheet ends at col 14)* |

Status is derived, not stored: no end date + start in future → **Scheduled**; no end date + started → **Running**; else **Complete**. A row is skipped unless it has a Product and a real `datetime` start. Product strings `Acrobat Reader`/`Acrobat Reduced Mode` are normalized to `Reader`/`Reduced Mode`.

### Fiscal calendar
`FY_STARTS` maps Adobe fiscal years to their start date (Saturday nearest Dec 1). `_extend_fy_starts()` auto-extends 3 FYs past today using a flat 364-day (52-week) roll. Adobe occasionally has 53-week fiscal years, so these auto-generated future starts drift and need manual correction roughly every ~5 years.

## Ingesting result screenshots

Result dashboards are captured as screenshots and turned into `Results` rows. The staging folder is **`results/`** (gitignored) and holds **both** the screenshots **and** the per-experiment workbooks (`RGSxxxx - … .xlsx`); processed images are moved to `results/processed/`.

**The key idea: cross-reference each screenshot to its RGS workbook.** A screenshot only shows summary metrics and the bare RGS id. The matching `RGSxxxx.xlsx` in `results/` supplies what the screenshot lacks — the **full test name + date range** (the summary tab's title cell) and exact **QGNARR/metrics** (the "Overall - Orders & ARR Scaled" block: `Units` and `Gross New ARR` rows → Control/Challenger/Lift%). Match on the **RGS id** found in the workbook's filename. The summary tab is named inconsistently (`Final Summary` / `Overall Summary` / `Summary`). Helpers live in `import_results.py`: `find_excel_for_rgs`, `read_excel_summary`, `enrich_from_excel` (excel is authoritative over screenshot OCR). **QGNARR mapping is a best guess** (Challenger Gross-New-ARR dollar figure) — verify against intent.

Two ways to extract — same target, same columns, same enrichment:

**A. In-session (preferred; no API key, no extra billing).** Claude Code reads the image with the Read tool, then enriches from the matching workbook (reuse the `import_results.py` helpers via a snippet) and writes the row via `openpyxl`. This is the right path when there's no `ANTHROPIC_API_KEY` (a Claude.ai/Claude Code subscription does **not** grant API access). Hard-won practices, in order:
1. **Close the workbook in Excel first**, then **back it up** — `openpyxl` rewrites the whole `.xlsx` on save and can drop charts/images on the `Gantt` sheet. Check for a `~$…xlsx` lock file to confirm Excel is closed; delete the backup only after verifying the result.
2. **Don't trust the product field** — dashboards often don't say Reader vs Reduced Mode. Infer from the matched workbook / an existing row for that RGS id, or ask; never just default it.
3. **Check for an existing row by RGS id before writing.** The same RGS id may already exist (a different surface, or an interim/partial row). If so, **update that row** rather than appending — and **confirm with the user before overwriting existing metrics** (e.g. interim → baked numbers).
4. Write per the **real sheet headers** (table above), set the same number formats as neighboring rows (`mm-dd-yy` dates, `0.00%` lifts), and **highlight new/updated rows yellow** (`PatternFill('solid', fgColor='FFFDE7')`) for review.
5. Rename the screenshot to the full test name and move it to `results/processed/`, then regenerate + `--deploy`.

**B. Scripted (`import_results.py`; needs `ANTHROPIC_API_KEY`).** Automates the same flow for batches: Claude vision extract → `enrich_from_excel` → append new rows (dedup on `(RGS id, product)`) → rename each processed image to its full test name in `results/processed/`. Use `--dry-run` to preview (extraction + enrichment, no write). Note the venv is Python 3.9, so the script relies on `from __future__ import annotations` for its `X | None` type hints.

## Seeding running experiments from Jira (`jira_sync.py`)

Populates the Results sheet with experiments that are *running* in Jira, so they appear on the roadmap before any results exist — writing **product, test name, start date only** (metrics/winner/end-date stay blank for later screenshot ingestion). No Jira MCP is connected, so it uses the **Jira REST API** with a token (env vars: `JIRA_BASE_URL` + either `JIRA_PAT` for Server/DC or `JIRA_EMAIL`/`JIRA_API_TOKEN` for Cloud). Config (JQL, which field is the start date, product-detection rules) is env-var/CONFIG-block driven; `--list-fields` dumps a sample issue's field ids to find custom fields. Dedups on `(RGS id, product)` via the shared `import_results` helpers, appends new rows highlighted yellow. Dependency-free (stdlib `urllib`). Meant to run manually and on a schedule.

## Gotchas

- **Output filename → `index.html` by design**: default `--output` is `Product Testing Roadmap.html` (gitignored on `main`); `--deploy` copies it to the publish clone as `index.html`. Locally generated `*.html` in this folder is scratch — don't commit it.
- `plan.json` is the planning persistence file (not config). It lives on `gh-pages` only, committed alongside `index.html` on deploy and updated directly by the browser via the Contents API. `--deploy` creates it as `[]` if missing but never clobbers browser-saved data.
- `~$...xlsx` is an Excel lock file (and a signal the workbook is open — its content names the lock owner). Ignore it; never write the `.xlsx` while it's present.
- Model ID is pinned in source: `import_results.py` uses `claude-opus-4-5`. Update it if migrating models.
