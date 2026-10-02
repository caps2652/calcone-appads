"""Configuration loading and defaults."""

from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from pathlib import Path


DEFAULTS = {
    "interval_seconds": 300,
    "output_dir": "screenshots",
    "image_format": "JPEG",
    "jpeg_quality": 60,
    "capture_all_monitors": True,
    "retention_days": 14,
    "show_startup_notice": True,
    "show_tray_icon": True,
    "notice_text": (
        "Bu bilgisayarda, güvenlik amacıyla ebeveyn tarafından "
        "ekran görüntüsü kaydı yapılmaktadır."
    ),
}


@dataclass
class Config:
    interval_seconds: int = DEFAULTS["interval_seconds"]
    output_dir: str = DEFAULTS["output_dir"]
    image_format: str = DEFAULTS["image_format"]
    jpeg_quality: int = DEFAULTS["jpeg_quality"]
    capture_all_monitors: bool = DEFAULTS["capture_all_monitors"]
    retention_days: int = DEFAULTS["retention_days"]
    show_startup_notice: bool = DEFAULTS["show_startup_notice"]
    show_tray_icon: bool = DEFAULTS["show_tray_icon"]
    notice_text: str = DEFAULTS["notice_text"]

    def validate(self) -> None:
        if self.interval_seconds < 5:
            raise ValueError("interval_seconds must be at least 5 seconds")
        fmt = self.image_format.upper()
        if fmt not in ("JPEG", "PNG"):
            raise ValueError("image_format must be 'JPEG' or 'PNG'")
        self.image_format = fmt
        if not (1 <= self.jpeg_quality <= 95):
            raise ValueError("jpeg_quality must be between 1 and 95")
        if self.retention_days < 0:
            raise ValueError("retention_days must be >= 0 (0 disables cleanup)")

    def resolved_output_dir(self, base: Path) -> Path:
        """Resolve output_dir relative to the config file's directory."""
        out = Path(self.output_dir)
        return out if out.is_absolute() else (base / out)

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2, ensure_ascii=False)


def load_config(path: Path | None) -> Config:
    """Load config from a JSON file, falling back to defaults.

    Unknown keys in the file are ignored; missing keys use defaults.
    """
    data = dict(DEFAULTS)
    if path is not None and path.exists():
        with path.open("r", encoding="utf-8") as fh:
            file_data = json.load(fh)
        for key in DEFAULTS:
            if key in file_data:
                data[key] = file_data[key]
    cfg = Config(**data)
    cfg.validate()
    return cfg
