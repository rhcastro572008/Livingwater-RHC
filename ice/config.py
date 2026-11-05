"""App settings. Everything can be overridden with environment variables (or a .env file)."""
from __future__ import annotations

import os
from pathlib import Path

try:  # optional: load a local .env file if python-dotenv is installed
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:  # pragma: no cover
    pass

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("ICE_DATA_DIR", BASE_DIR / "data"))

TRACKER_PATH = Path(os.environ.get("TRACKER_PATH", DATA_DIR / "Tracker.xlsx"))
REPORTS_DIR = DATA_DIR / "reports"      # generated daily Ice Monitoring Reports (.xlsx)
RAW_DIR = DATA_DIR / "raw"              # original whiteboard photos, one per report date
UPLOADS_DIR = DATA_DIR / "uploads"      # photos waiting for review
BACKUPS_DIR = DATA_DIR / "backups"      # Tracker.xlsx copies taken before every write

# Rider order is the order used in the daily report.
RIDERS = [
    r.strip()
    for r in os.environ.get("RIDERS", "Amped,Khurtz,Ram,Mark,Mike,Jomar,Pickup").split(",")
    if r.strip()
]

# Known short forms written on the board -> Client List spelling.
CLIENT_ALIASES = {
    "LITO'S": "LITO FOODHAUSE",
    "LITOS": "LITO FOODHAUSE",
    "JOY": "JOY STORE",
    "L300-180": "L300",
}

# Dashboard: clients with no order in this many days are listed as "not ordering".
INACTIVE_DAYS = int(os.environ.get("INACTIVE_DAYS", "5"))

# Claude vision model used to read the whiteboard photo.
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-opus-5-5")

SECRET_KEY = os.environ.get("SECRET_KEY", "change-me-in-production")
# Optional shared password for the whole site (HTTP basic auth, any username).
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "20"))


def ensure_dirs() -> None:
    for d in (DATA_DIR, REPORTS_DIR, RAW_DIR, UPLOADS_DIR, BACKUPS_DIR, TRACKER_PATH.parent):
        d.mkdir(parents=True, exist_ok=True)
