#!/usr/bin/env python3
"""
jira_sync.py — Pull *running* experiments from Jira into the roadmap.

Queries Jira via JQL, maps each issue to (product, test name, start date), and
appends new rows to the Results sheet — dedup'd on (RGS id, product) against what's
already there. Metrics/winner/end-date are left blank on purpose, to be filled later
by screenshot + xlsx ingestion (import_results.py). New rows are highlighted yellow.

There is no Jira MCP connected, so this talks to the Jira REST API directly.

AUTH — set as environment variables (nothing secret is stored in this file):
  Atlassian Cloud:     JIRA_BASE_URL, JIRA_EMAIL, JIRA_API_TOKEN
  Server / DataCenter: JIRA_BASE_URL, JIRA_PAT        (Personal Access Token, sent as Bearer)

CONFIG — set as env vars or edit the CONFIG block below:
  JIRA_JQL    the query that selects your running experiments
  field map   which Jira field holds the start date (run --list-fields to discover ids)

USAGE:
  python jira_sync.py --list-fields     # dump one matching issue's fields → find custom field ids
  python jira_sync.py --dry-run         # show what would be added, write nothing
  python jira_sync.py                   # append new running experiments to the Results sheet
  python jira_sync.py --jql 'project = RGS AND status = Running'   # override the JQL once
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

# Reuse the roadmap helpers (Excel columns, RGS-id dedup, date parsing) so the layout
# lives in one place. import_results imports anthropic lazily, so this is dependency-free.
from import_results import (
    DEFAULT_EXCEL, rgs_id, dedup_key, parse_date, append_row, get_existing_keys, load_workbook,
)

# ── CONFIG — adjust to match your Jira ────────────────────────────────────────────
JIRA_BASE_URL = os.environ.get('JIRA_BASE_URL', '')   # e.g. https://jira.corp.adobe.com

# The query that selects experiments that are currently running. Override with --jql or
# the JIRA_JQL env var. TODO: confirm project key / issue type / status name.
DEFAULT_JQL = os.environ.get(
    'JIRA_JQL',
    'project = RGS AND status = "Running" ORDER BY created ASC',
)

# Which Jira field supplies each roadmap value. 'summary' / 'created' are standard;
# custom fields look like 'customfield_12345' (use --list-fields to find them).
TEST_NAME_FIELD  = os.environ.get('JIRA_TEST_NAME_FIELD', 'summary')
START_DATE_FIELD = os.environ.get('JIRA_START_DATE_FIELD', 'created')  # TODO: real "Start date" field?

# How to decide Reader vs Reduced Mode: which issue attribute to inspect, and the
# substrings that map to each product. Checks labels, components, then summary text.
PRODUCT_SOURCES = ('labels', 'components', 'summary')
PRODUCT_RULES = [          # (substring to look for, value to write) — first match wins
    ('reduced',     'Acrobat Reduced Mode'),
    ('reader',      'Acrobat Reader'),
]
DEFAULT_PRODUCT = 'Acrobat Reader'   # fallback when nothing matches
# ──────────────────────────────────────────────────────────────────────────────────


def auth_header() -> dict:
    """Build the auth header from env vars: Bearer PAT (Server/DC) or Basic email:token (Cloud)."""
    pat = os.environ.get('JIRA_PAT')
    if pat:
        return {'Authorization': f'Bearer {pat}'}
    email = os.environ.get('JIRA_EMAIL')
    token = os.environ.get('JIRA_API_TOKEN')
    if email and token:
        basic = base64.b64encode(f'{email}:{token}'.encode()).decode()
        return {'Authorization': f'Basic {basic}'}
    sys.exit("ERROR: set JIRA_PAT (Server/DC) or JIRA_EMAIL + JIRA_API_TOKEN (Cloud).")


def jira_request(path: str, payload: dict) -> dict:
    """POST to the Jira REST API and return parsed JSON."""
    if not JIRA_BASE_URL:
        sys.exit("ERROR: set JIRA_BASE_URL (e.g. https://jira.corp.adobe.com).")
    url = JIRA_BASE_URL.rstrip('/') + path
    data = json.dumps(payload).encode()
    headers = {'Content-Type': 'application/json', 'Accept': 'application/json', **auth_header()}
    req = urllib.request.Request(url, data=data, headers=headers, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors='replace')[:500]
        sys.exit(f"ERROR: Jira returned {e.code} for {path}\n  {body}")
    except urllib.error.URLError as e:
        sys.exit(f"ERROR: could not reach Jira at {JIRA_BASE_URL}: {e.reason}")


def search_issues(jql: str, fields: list[str] | None = None) -> list[dict]:
    """Run a JQL search, paging through all results."""
    issues, start = [], 0
    while True:
        page = jira_request('/rest/api/2/search', {
            'jql': jql, 'startAt': start, 'maxResults': 100,
            'fields': fields if fields is not None else ['summary', 'labels', 'components',
                                                          'created', START_DATE_FIELD],
        })
        batch = page.get('issues', [])
        issues.extend(batch)
        start += len(batch)
        if start >= page.get('total', 0) or not batch:
            return issues


def map_product(fields: dict) -> str:
    """Infer Reader vs Reduced Mode from the configured issue attributes."""
    hay = []
    for src in PRODUCT_SOURCES:
        v = fields.get(src)
        if src == 'labels' and isinstance(v, list):
            hay += [str(x) for x in v]
        elif src == 'components' and isinstance(v, list):
            hay += [c.get('name', '') for c in v]
        elif isinstance(v, str):
            hay.append(v)
    blob = ' '.join(hay).lower()
    for needle, product in PRODUCT_RULES:
        if needle in blob:
            return product
    return DEFAULT_PRODUCT


def normalize_date(val) -> str | None:
    """Jira dates may be 'YYYY-MM-DD' or full ISO timestamps; return YYYY-MM-DD."""
    if not val:
        return None
    s = str(val)
    return s.split('T')[0] if 'T' in s else s


def map_issue(issue: dict) -> dict:
    """Turn a Jira issue into a roadmap row: product, test name, start date only."""
    fields = issue.get('fields', {})
    raw_name = fields.get(TEST_NAME_FIELD) or fields.get('summary') or issue.get('key', '')
    return {
        'product': map_product(fields),
        'test_name': str(raw_name).strip(),
        'start_date': normalize_date(fields.get(START_DATE_FIELD) or fields.get('created')),
        '_key': issue.get('key'),
    }


def cmd_list_fields(jql: str):
    """Print field ids + values for one matching issue, to help map custom fields."""
    issues = search_issues(jql, fields=['*all'])[:1]
    if not issues:
        print(f"No issues matched: {jql}")
        return
    # field id -> human name
    names = {f['id']: f.get('name', f['id']) for f in jira_request_get('/rest/api/2/field')}
    fields = issues[0].get('fields', {})
    print(f"Sample issue {issues[0].get('key')} — non-empty fields:\n")
    for fid, val in sorted(fields.items()):
        if val in (None, '', [], {}):
            continue
        preview = str(val)
        if len(preview) > 70:
            preview = preview[:70] + '…'
        print(f"  {fid:24} {names.get(fid, '')!r:32} = {preview}")


def jira_request_get(path: str) -> list:
    """GET helper (used for /field metadata)."""
    if not JIRA_BASE_URL:
        sys.exit("ERROR: set JIRA_BASE_URL.")
    req = urllib.request.Request(JIRA_BASE_URL.rstrip('/') + path,
                                 headers={'Accept': 'application/json', **auth_header()})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())


def main():
    ap = argparse.ArgumentParser(description='Sync running experiments from Jira into the roadmap Excel.')
    ap.add_argument('--excel', '-e', default=str(DEFAULT_EXCEL), help='Path to the roadmap Excel file')
    ap.add_argument('--jql', default=DEFAULT_JQL, help='Override the JQL for this run')
    ap.add_argument('--dry-run', action='store_true', help='Show what would be added; write nothing')
    ap.add_argument('--list-fields', action='store_true',
                    help='Dump one matching issue\'s fields to discover custom field ids')
    args = ap.parse_args()

    if args.list_fields:
        cmd_list_fields(args.jql)
        return

    print(f"Querying Jira: {args.jql}")
    issues = search_issues(args.jql)
    rows = [map_issue(i) for i in issues]
    print(f"  {len(rows)} running experiment(s) returned")

    excel = os.path.expanduser(args.excel)
    if not os.path.exists(excel):
        sys.exit(f"ERROR: roadmap Excel not found: {excel}")
    wb = load_workbook(excel)
    ws = wb['Results']
    existing = get_existing_keys(ws)

    to_add = []
    for r in rows:
        if not r['test_name'] or not r['start_date']:
            print(f"  SKIP (missing name/start): {r.get('_key')}")
            continue
        if dedup_key(r['test_name'], r['product']) in existing:
            continue
        to_add.append(r)

    print(f"  {len(to_add)} new (not already in the sheet)")
    for r in to_add:
        print(f"    + [{r['_key']}] {r['product']} | {r['test_name']} | {r['start_date']}")

    if args.dry_run:
        print("\n── DRY RUN — nothing written ──")
        return
    if not to_add:
        print("Nothing to add.")
        return

    next_row = ws.max_row + 1
    for r in to_add:
        append_row(ws, r, next_row)
        existing.add(dedup_key(r['test_name'], r['product']))
        next_row += 1
    wb.save(excel)
    print(f"\n✓ Added {len(to_add)} row(s) to {excel} (highlighted yellow — review before deploying)")
    print("Next: python generate_gantt.py --deploy")


if __name__ == '__main__':
    main()
