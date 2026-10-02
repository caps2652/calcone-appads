"""Command-line interface."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import Config, load_config
from .capture import capture_once
from .gallery import build_gallery
from .monitor import run_monitor


def _base_dir(config_path: Path | None) -> Path:
    return config_path.resolve().parent if config_path else Path.cwd()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="screen_monitor",
        description="Şeffaf ebeveyn ekran izleme aracı (yerel, periyodik ekran görüntüsü).",
    )
    parser.add_argument(
        "-c", "--config", type=Path, default=Path("config.json"),
        help="JSON yapılandırma dosyası yolu (varsayılan: config.json).",
    )
    sub = parser.add_subparsers(dest="command")

    sub.add_parser("run", help="İzlemeyi sürekli çalıştır (varsayılan).")
    sub.add_parser("once", help="Tek bir ekran görüntüsü al ve çık.")
    sub.add_parser("gallery", help="HTML galeriyi yeniden oluştur ve çık.")

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config_path = args.config if args.config and args.config.exists() else None
    if args.config and not args.config.exists():
        print(
            f"[ekran-izleme] Not: '{args.config}' bulunamadı, varsayılan ayarlar "
            f"kullanılıyor. (config.example.json dosyasını kopyalayabilirsiniz.)"
        )

    try:
        cfg: Config = load_config(config_path)
    except (ValueError, OSError) as exc:
        print(f"[ekran-izleme] Yapılandırma hatası: {exc}", file=sys.stderr)
        return 2

    output_dir = cfg.resolved_output_dir(_base_dir(config_path))
    command = args.command or "run"

    if command == "once":
        path = capture_once(cfg, output_dir)
        build_gallery(output_dir)
        print(f"[ekran-izleme] Kaydedildi: {path}")
        return 0

    if command == "gallery":
        path = build_gallery(output_dir)
        print(f"[ekran-izleme] Galeri oluşturuldu: {path}")
        return 0

    # Default: run the monitor loop.
    try:
        run_monitor(cfg, output_dir)
    except KeyboardInterrupt:
        print("\n[ekran-izleme] Durduruldu.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
