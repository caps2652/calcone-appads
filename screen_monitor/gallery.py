"""Generate a simple local HTML gallery for reviewing screenshots."""

from __future__ import annotations

import html
from pathlib import Path

IMAGE_EXTS = {".jpg", ".jpeg", ".png"}


def _collect(output_dir: Path) -> dict[str, list[Path]]:
    """Return {day_folder_name: [image_paths...]} sorted newest first."""
    by_day: dict[str, list[Path]] = {}
    if not output_dir.exists():
        return by_day
    for day_dir in sorted(output_dir.iterdir(), reverse=True):
        if not day_dir.is_dir() or day_dir.name.startswith("."):
            continue
        images = sorted(
            (p for p in day_dir.iterdir() if p.suffix.lower() in IMAGE_EXTS),
            reverse=True,
        )
        if images:
            by_day[day_dir.name] = images
    return by_day


def build_gallery(output_dir: Path) -> Path:
    """Write gallery.html into output_dir and return its path."""
    by_day = _collect(output_dir)
    gallery_path = output_dir / "gallery.html"

    parts = [
        "<!doctype html>",
        '<html lang="tr"><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        "<title>Ekran Görüntüleri</title>",
        "<style>",
        "body{font-family:system-ui,sans-serif;margin:0;background:#0f1115;color:#e6e6e6}",
        "header{padding:16px 20px;background:#1a1d24;position:sticky;top:0}",
        "h1{font-size:18px;margin:0}",
        "h2{font-size:15px;margin:24px 20px 8px;color:#9fb4d8}",
        ".grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));"
        "gap:10px;padding:0 20px}",
        ".card{background:#1a1d24;border-radius:8px;overflow:hidden}",
        ".card img{width:100%;height:140px;object-fit:cover;display:block}",
        ".card a{color:#9fb4d8;text-decoration:none;font-size:12px;display:block;"
        "padding:6px 8px}",
        ".empty{padding:40px 20px;color:#888}",
        "</style></head><body>",
        "<header><h1>Ekran Görüntüleri</h1></header>",
    ]

    total = sum(len(v) for v in by_day.values())
    if total == 0:
        parts.append('<p class="empty">Henüz kayıtlı ekran görüntüsü yok.</p>')
    else:
        for day, images in by_day.items():
            parts.append(f"<h2>{html.escape(day)} &middot; {len(images)} görüntü</h2>")
            parts.append('<div class="grid">')
            for img in images:
                rel = f"{day}/{img.name}"
                label = html.escape(img.stem.replace("-", ":"))
                parts.append(
                    f'<div class="card"><a href="{rel}" target="_blank">'
                    f'<img loading="lazy" src="{rel}" alt="{label}">'
                    f"<span>{label}</span></a></div>"
                )
            parts.append("</div>")

    parts.append("</body></html>")
    gallery_path.write_text("\n".join(parts), encoding="utf-8")
    return gallery_path
