"""Visible notifications that keep the monitoring transparent.

A startup notice and an optional system tray icon make it clear to anyone
using the computer that screen captures are being recorded. This is a
deliberate design choice: the tool is a parental-control aid, not covert
surveillance.
"""

from __future__ import annotations

import threading


def show_startup_notice(text: str) -> None:
    """Show a one-time, non-blocking notice that monitoring is active.

    Tries a native Windows message box, then falls back to tkinter, and
    finally to stdout. Failures are swallowed so they never stop monitoring.
    """

    def _worker() -> None:
        # 1) Native Windows message box.
        try:
            import ctypes  # noqa: PLC0415

            MB_OK = 0x0
            MB_ICONINFORMATION = 0x40
            ctypes.windll.user32.MessageBoxW(
                0, text, "Ekran İzleme Etkin", MB_OK | MB_ICONINFORMATION
            )
            return
        except Exception:
            pass

        # 2) Cross-platform tkinter fallback.
        try:
            import tkinter  # noqa: PLC0415
            from tkinter import messagebox  # noqa: PLC0415

            root = tkinter.Tk()
            root.withdraw()
            messagebox.showinfo("Ekran İzleme Etkin", text)
            root.destroy()
            return
        except Exception:
            pass

        # 3) Last resort.
        print(f"[ekran-izleme] {text}")

    threading.Thread(target=_worker, daemon=True).start()


def run_tray_icon(stop_event: threading.Event, output_dir, on_quit) -> bool:
    """Run a system tray icon until `stop_event` is set.

    Returns True if the tray ran, False if pystray/Pillow were unavailable.
    This blocks the calling thread, so run it from the main thread.
    """
    try:
        import pystray  # noqa: PLC0415
        from PIL import Image, ImageDraw  # noqa: PLC0415
    except Exception:
        return False

    # Simple generated icon: a blue circle with an eye-like dot.
    size = 64
    image = Image.new("RGB", (size, size), (33, 94, 171))
    draw = ImageDraw.Draw(image)
    draw.ellipse((14, 20, 50, 44), fill=(255, 255, 255))
    draw.ellipse((27, 26, 37, 38), fill=(33, 94, 171))

    def _quit(icon, _item):
        on_quit()
        icon.stop()

    menu = pystray.Menu(
        pystray.MenuItem(f"Kayıt klasörü: {output_dir}", None, enabled=False),
        pystray.MenuItem("Ekran izlemeyi durdur", _quit),
    )
    icon = pystray.Icon("ekran-izleme", image, "Ekran İzleme (aktif)", menu)

    # Stop the icon when the stop_event fires from elsewhere.
    def _watch():
        stop_event.wait()
        try:
            icon.stop()
        except Exception:
            pass

    threading.Thread(target=_watch, daemon=True).start()
    icon.run()
    return True
