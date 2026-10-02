"""Screenshot capture using the `mss` library."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import mss
from PIL import Image

from .config import Config


def _timestamp_name(now: datetime, ext: str) -> str:
    return now.strftime("%H-%M-%S") + f".{ext}"


def _day_dir(base: Path, now: datetime) -> Path:
    day = base / now.strftime("%Y-%m-%d")
    day.mkdir(parents=True, exist_ok=True)
    return day


def capture_once(cfg: Config, output_dir: Path, now: datetime | None = None) -> Path:
    """Capture a single screenshot and save it under a dated subfolder.

    Returns the path of the saved image.
    """
    now = now or datetime.now()
    output_dir.mkdir(parents=True, exist_ok=True)

    with mss.mss() as sct:
        # monitors[0] is the virtual bounding box across all displays;
        # monitors[1] is the primary display.
        monitor = sct.monitors[0] if cfg.capture_all_monitors else sct.monitors[1]
        shot = sct.grab(monitor)
        img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

    ext = "jpg" if cfg.image_format == "JPEG" else "png"
    target = _day_dir(output_dir, now) / _timestamp_name(now, ext)

    if cfg.image_format == "JPEG":
        img.save(target, "JPEG", quality=cfg.jpeg_quality, optimize=True)
    else:
        img.save(target, "PNG", optimize=True)

    return target
