"""Retention / cleanup of stored screenshots."""

from __future__ import annotations

import shutil
from datetime import datetime, timedelta
from pathlib import Path


def cleanup_old(output_dir: Path, retention_days: int, now: datetime | None = None) -> int:
    """Delete dated subfolders older than `retention_days`.

    Date folders are named YYYY-MM-DD. Returns the number of folders removed.
    A retention_days of 0 disables cleanup.
    """
    if retention_days <= 0 or not output_dir.exists():
        return 0

    now = now or datetime.now()
    cutoff = (now - timedelta(days=retention_days)).date()
    removed = 0

    for child in output_dir.iterdir():
        if not child.is_dir():
            continue
        try:
            folder_date = datetime.strptime(child.name, "%Y-%m-%d").date()
        except ValueError:
            # Not a dated folder (e.g. the generated gallery); leave it alone.
            continue
        if folder_date < cutoff:
            shutil.rmtree(child, ignore_errors=True)
            removed += 1

    return removed
