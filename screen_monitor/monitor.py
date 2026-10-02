"""Main monitoring loop."""

from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path

from .capture import capture_once
from .config import Config
from .gallery import build_gallery
from .notice import run_tray_icon, show_startup_notice
from .storage import cleanup_old


def run_monitor(cfg: Config, output_dir: Path) -> None:
    """Run the capture loop until interrupted (Ctrl+C) or the tray quits."""
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"[ekran-izleme] Başladı. Aralık: {cfg.interval_seconds}s")
    print(f"[ekran-izleme] Kayıt klasörü: {output_dir}")

    if cfg.show_startup_notice:
        show_startup_notice(cfg.notice_text)

    stop_event = threading.Event()

    def _capture_loop() -> None:
        # Capture immediately, then on each interval.
        while not stop_event.is_set():
            now = datetime.now()
            try:
                path = capture_once(cfg, output_dir, now)
                print(f"[ekran-izleme] {now:%H:%M:%S} kaydedildi: {path.name}")
                removed = cleanup_old(output_dir, cfg.retention_days, now)
                if removed:
                    print(f"[ekran-izleme] {removed} eski klasör temizlendi")
                build_gallery(output_dir)
            except Exception as exc:  # keep the loop alive on transient errors
                print(f"[ekran-izleme] Uyarı: kayıt alınamadı: {exc}")
            stop_event.wait(cfg.interval_seconds)

    worker = threading.Thread(target=_capture_loop, daemon=True)
    worker.start()

    # The tray icon (if enabled and available) must run on the main thread.
    tray_ran = False
    if cfg.show_tray_icon:
        tray_ran = run_tray_icon(stop_event, output_dir, on_quit=stop_event.set)

    if not tray_ran:
        # No tray: block on the worker and honour Ctrl+C.
        try:
            while worker.is_alive():
                worker.join(timeout=0.5)
        except KeyboardInterrupt:
            pass

    stop_event.set()
    worker.join(timeout=cfg.interval_seconds + 5)
    build_gallery(output_dir)
    print("[ekran-izleme] Durduruldu.")
