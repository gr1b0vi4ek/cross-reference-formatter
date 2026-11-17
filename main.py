"""Точка входа CRF/ФПС."""

import ctypes
import sys


def _set_dpi_awareness() -> None:
    if sys.platform != 'win32':
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main() -> None:
    _set_dpi_awareness()
    from src.gui import CrossReferencerApp
    CrossReferencerApp().mainloop()


if __name__ == '__main__':
    main()