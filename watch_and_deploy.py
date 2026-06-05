#!/usr/bin/env python3
"""
watch_and_deploy.py — Watches the Excel file for changes and auto-deploys
the Gantt to GitHub Pages.

USAGE:
  python3 watch_and_deploy.py              # run in foreground
  python3 watch_and_deploy.py &            # run in background
  nohup python3 watch_and_deploy.py &      # persist after terminal closes

STOP:
  kill $(cat ~/.gantt-watcher.pid)         # stop the watcher

SETUP (one-time):
  pip3 install watchdog
"""

import os
import sys
import time
import subprocess
import signal
from pathlib import Path
from datetime import datetime

# ── Configuration ────────────────────────────────────────────
ONEDRIVE_FOLDER = Path.home() / 'Library' / 'CloudStorage' / 'OneDrive-Adobe' / 'Reader & Reduced Mode - Infra Testing'
EXCEL_FILE = ONEDRIVE_FOLDER / 'Reader & Reduced Mode Testing Roadmap.xlsx'
SCRIPT_FILE = ONEDRIVE_FOLDER / 'generate_gantt.py'
REPO_PATH = Path.home() / 'infra-testing-roadmap'
PID_FILE = Path.home() / '.gantt-watcher.pid'
COOLDOWN = 10  # seconds to wait after a change before deploying (debounce)
# ─────────────────────────────────────────────────────────────

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    print("ERROR: watchdog is required. Install it with: pip3 install watchdog")
    sys.exit(1)


def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f"[{ts}] {msg}", flush=True)


def deploy():
    """Run generate_gantt.py --deploy."""
    log("Excel changed — regenerating and deploying...")
    try:
        result = subprocess.run(
            [sys.executable, str(SCRIPT_FILE), '--deploy', '--repo', str(REPO_PATH)],
            capture_output=True, text=True, timeout=120
        )
        # Print output
        for line in result.stdout.strip().splitlines():
            log(f"  {line}")
        if result.returncode != 0:
            log(f"  ERROR: {result.stderr.strip()}")
        else:
            log("Deploy complete.")
    except subprocess.TimeoutExpired:
        log("  ERROR: Deploy timed out after 120s")
    except Exception as e:
        log(f"  ERROR: {e}")


class ExcelChangeHandler(FileSystemEventHandler):
    def __init__(self):
        self.last_trigger = 0

    def on_modified(self, event):
        # Only react to the Excel file
        if not event.src_path.endswith('.xlsx'):
            return
        if Path(event.src_path).name != EXCEL_FILE.name:
            return
        # Debounce: Excel saves trigger multiple modification events
        now = time.time()
        if now - self.last_trigger < COOLDOWN:
            return
        self.last_trigger = now
        # Small delay to let Excel finish writing
        time.sleep(2)
        deploy()


def cleanup(signum, frame):
    log("Stopping watcher...")
    if PID_FILE.exists():
        PID_FILE.unlink()
    sys.exit(0)


def main():
    # Validate paths
    if not EXCEL_FILE.exists():
        log(f"ERROR: Excel file not found: {EXCEL_FILE}")
        sys.exit(1)
    if not SCRIPT_FILE.exists():
        log(f"ERROR: Script not found: {SCRIPT_FILE}")
        sys.exit(1)
    if not (REPO_PATH / '.git').exists():
        log(f"ERROR: Git repo not found at: {REPO_PATH}")
        sys.exit(1)

    # Write PID file for easy stopping
    PID_FILE.write_text(str(os.getpid()))
    signal.signal(signal.SIGTERM, cleanup)
    signal.signal(signal.SIGINT, cleanup)

    log(f"Watching: {EXCEL_FILE.name}")
    log(f"Repo:     {REPO_PATH}")
    log(f"PID:      {os.getpid()} (saved to {PID_FILE})")
    log(f"Cooldown: {COOLDOWN}s between deploys")
    log("Waiting for changes... (Ctrl+C to stop)")

    handler = ExcelChangeHandler()
    observer = Observer()
    observer.schedule(handler, str(ONEDRIVE_FOLDER), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    cleanup(None, None)


if __name__ == '__main__':
    main()
