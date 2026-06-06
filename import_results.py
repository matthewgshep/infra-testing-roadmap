#!/usr/bin/env python3
"""
import_results.py — Reads experiment result screenshots from a folder,
uses Claude vision to extract test data, and appends new rows to the
Results sheet in the Testing Roadmap Excel file.

SETUP:
  pip install openpyxl anthropic pillow

USAGE:
  python import_results.py                        # uses default paths
  python import_results.py --folder ~/Desktop/screenshots
  python import_results.py --folder ~/Desktop/screenshots --dry-run

WORKFLOW:
  1. Drop screenshot images into the screenshots folder
  2. Run this script
  3. Review the newly appended rows in Excel
  4. Run generate_gantt.py to regenerate the HTML
"""

from __future__ import annotations  # allow `X | None` annotations on Python 3.9

import argparse
import base64
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import anthropic
except ImportError:
    print("ERROR: anthropic package required. Run: pip install anthropic")
    sys.exit(1)

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    print("ERROR: openpyxl required. Run: pip install openpyxl")
    sys.exit(1)

ONEDRIVE = Path.home() / 'Library' / 'CloudStorage' / 'OneDrive-Adobe' / 'Reader & Reduced Mode - Infra Testing'
DEFAULT_EXCEL  = ONEDRIVE / 'Reader & Reduced Mode Testing Roadmap.xlsx'
DEFAULT_FOLDER = ONEDRIVE / 'screenshots'

IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.webp', '.gif'}

EXTRACT_PROMPT = """You are analyzing a screenshot of an A/B experiment results dashboard for Adobe Acrobat (Reader or Reduced Mode product testing).

Extract ALL experiment results visible in the screenshot. Return a JSON array — one object per experiment row — with these exact keys:

- "product": "Acrobat Reader" or "Acrobat Reduced Mode" (infer from context; default to "Acrobat Reader" if unclear)
- "test_name": full test name including RGS ID, e.g. "RGS0480 - Export PDF Direct Upsell Modal - Reader - Multi - WW - EN"
- "start_date": "YYYY-MM-DD" or null
- "end_date": "YYYY-MM-DD" or null (null if still running)
- "qgnarr": numeric QGNARR value (integer) or null
- "gnarr_lift": decimal lift e.g. 0.022 for 2.2%, or null
- "ctr": decimal CTR lift or null
- "units_pct": decimal units % lift or null
- "winner": one of "Challenger 1", "Challenger 2", "Control", "Rollout", or null
- "details": decision notes or notable commentary as a string, or null
- "gds_url": full URL if a GDS/dashboard link is visible, or null

Return ONLY a valid JSON array with no markdown fences or explanation."""


def encode_image(path: Path) -> tuple[str, str]:
    ext = path.suffix.lower()
    media_map = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                 '.webp': 'image/webp', '.gif': 'image/gif'}
    media_type = media_map.get(ext, 'image/png')
    with open(path, 'rb') as f:
        return base64.standard_b64encode(f.read()).decode(), media_type


def extract_from_screenshot(client: anthropic.Anthropic, image_path: Path) -> list[dict]:
    print(f"  Extracting: {image_path.name}")
    b64, media_type = encode_image(image_path)
    msg = client.messages.create(
        model='claude-opus-4-5',
        max_tokens=2000,
        messages=[{
            'role': 'user',
            'content': [
                {'type': 'image', 'source': {'type': 'base64', 'media_type': media_type, 'data': b64}},
                {'type': 'text', 'text': EXTRACT_PROMPT}
            ]
        }]
    )
    raw = msg.content[0].text.strip()
    raw = raw.replace('```json', '').replace('```', '').strip()
    parsed = json.loads(raw)
    return parsed if isinstance(parsed, list) else [parsed]


def rgs_id(name) -> str | None:
    """Extract the RGS identifier (e.g. 'RGS0469') from a test name, if present."""
    m = re.search(r'RGS\d+', str(name or ''), re.IGNORECASE)
    return m.group(0).upper() if m else None


def dedup_key(test_name, product):
    """Key a row by (RGS id, product) so the same experiment is caught even when the
    rest of the test name differs, while genuine per-surface variants stay distinct.
    Falls back to the full name when there's no RGS id."""
    rid = rgs_id(test_name)
    prod = str(product or '').strip()
    return (rid, prod) if rid else (None, str(test_name or '').strip())


def get_existing_keys(ws) -> set:
    keys = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[2]:  # Test Name is column C (index 2); Product is column B (index 1)
            keys.add(dedup_key(row[2], row[1]))
    return keys


def parse_date(val) -> datetime | None:
    if not val:
        return None
    if isinstance(val, datetime):
        return val
    for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(str(val).strip(), fmt)
        except ValueError:
            continue
    return None


def append_row(ws, row: dict, row_num: int):
    # Match column order: A(blank), B(Product), C(Test Name), D(Start), E(End),
    # F(QGNARR), G(GNARR Lift), H(CTR), I(Units %), J(Winner), K(GDS Links),
    # L(Decision/Details), M(Result contrary), N(PM Commentary)
    ws.cell(row_num, 1).value = None
    ws.cell(row_num, 2).value = row.get('product') or ''
    ws.cell(row_num, 3).value = row.get('test_name') or ''
    ws.cell(row_num, 4).value = parse_date(row.get('start_date'))
    ws.cell(row_num, 5).value = parse_date(row.get('end_date'))

    qgnarr = row.get('qgnarr')
    ws.cell(row_num, 6).value = int(qgnarr) if qgnarr is not None else None

    for col, key in [(7, 'gnarr_lift'), (8, 'ctr'), (9, 'units_pct')]:
        v = row.get(key)
        ws.cell(row_num, col).value = float(v) if v is not None else None

    ws.cell(row_num, 10).value = row.get('winner') or None

    gds_url = row.get('gds_url')
    if gds_url:
        cell = ws.cell(row_num, 11)
        cell.value = 'Link'
        cell.hyperlink = gds_url
        cell.font = Font(color='2E75B6', underline='single')
    else:
        ws.cell(row_num, 11).value = None

    ws.cell(row_num, 12).value = row.get('details') or None

    # Date formatting
    for col in (4, 5):
        c = ws.cell(row_num, col)
        if c.value:
            c.number_format = 'YYYY-MM-DD'

    # Highlight new row in light yellow so it's easy to spot
    fill = PatternFill('solid', fgColor='FFFDE7')
    for col in range(1, 15):
        ws.cell(row_num, col).fill = fill


def main():
    parser = argparse.ArgumentParser(description='Import experiment screenshots into the Testing Roadmap Excel.')
    parser.add_argument('--folder', '-f', default=str(DEFAULT_FOLDER),
                        help='Folder containing screenshot images')
    parser.add_argument('--excel', '-e', default=str(DEFAULT_EXCEL),
                        help='Path to the Testing Roadmap Excel file')
    parser.add_argument('--dry-run', action='store_true',
                        help='Extract and print results without writing to Excel')
    parser.add_argument('--api-key', default=os.environ.get('ANTHROPIC_API_KEY'),
                        help='Anthropic API key (or set ANTHROPIC_API_KEY env var)')
    args = parser.parse_args()

    # Resolve paths
    folder = Path(args.folder).expanduser().resolve()
    excel  = Path(args.excel).expanduser().resolve()

    # Validate
    if not folder.exists():
        print(f"Screenshots folder not found: {folder}")
        print(f"Creating it for you...")
        folder.mkdir(parents=True, exist_ok=True)
        print(f"Drop your screenshots into:\n  {folder}\nThen run this script again.")
        sys.exit(0)

    images = sorted([p for p in folder.iterdir() if p.suffix.lower() in IMAGE_EXTS])
    if not images:
        print(f"No images found in: {folder}")
        print(f"Supported formats: {', '.join(IMAGE_EXTS)}")
        sys.exit(0)

    print(f"Found {len(images)} image(s) in {folder}")

    if not args.api_key:
        print("\nERROR: Anthropic API key required.")
        print("  Option 1: export ANTHROPIC_API_KEY=sk-ant-...")
        print("  Option 2: python import_results.py --api-key sk-ant-...")
        print("\nGet a key at: https://console.anthropic.com/settings/keys")
        sys.exit(1)

    if not excel.exists():
        print(f"ERROR: Excel file not found: {excel}")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=args.api_key)

    # Extract from all screenshots
    all_rows = []
    errors = []
    for img in images:
        try:
            rows = extract_from_screenshot(client, img)
            print(f"    → {len(rows)} row(s) extracted")
            for r in rows:
                r['_source'] = img.name
            all_rows.extend(rows)
        except Exception as e:
            print(f"    ✗ Error: {e}")
            errors.append(img.name)

    if not all_rows:
        print("\nNo data extracted.")
        sys.exit(0)

    print(f"\nTotal: {len(all_rows)} row(s) extracted from {len(images) - len(errors)} image(s)")

    if args.dry_run:
        print("\n── DRY RUN — not writing to Excel ──")
        for r in all_rows:
            print(f"  [{r.get('_source')}] {r.get('test_name')} | {r.get('start_date')} → {r.get('end_date')} | winner: {r.get('winner')} | GNARR: {r.get('gnarr_lift')}")
        sys.exit(0)

    # Write to Excel
    print(f"\nOpening: {excel}")
    wb = load_workbook(excel)
    ws = wb['Results']

    existing = get_existing_keys(ws)
    next_row = ws.max_row + 1
    added = 0
    skipped = 0

    for row in all_rows:
        name = (row.get('test_name') or '').strip()
        key = dedup_key(name, row.get('product'))
        if key in existing:
            print(f"  SKIP (already exists): {name}")
            skipped += 1
            continue
        append_row(ws, row, next_row)
        print(f"  ADDED row {next_row}: {name}")
        existing.add(key)
        next_row += 1
        added += 1

    if added == 0:
        print("\nAll rows already exist in the sheet — nothing to add.")
        sys.exit(0)

    wb.save(excel)
    print(f"\n✓ Saved {added} new row(s) to: {excel}")
    print(f"  (highlighted in yellow — review before regenerating the Gantt)")
    if skipped:
        print(f"  {skipped} duplicate(s) skipped")
    print(f"\nNext step: python generate_gantt.py --deploy")

    # Move processed images to a 'processed' subfolder
    processed_dir = folder / 'processed'
    processed_dir.mkdir(exist_ok=True)
    for img in images:
        if img.name not in errors:
            img.rename(processed_dir / img.name)
    print(f"  Screenshots moved to: {processed_dir}")


if __name__ == '__main__':
    main()
