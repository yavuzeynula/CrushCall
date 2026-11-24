"""Uygulama girişi (src/ içinde, güncellenebilir): pencere + arkada güncelleme kontrolü.

Yeni sürüm varsa kullanım durur ve güncellenir (diğer programlardaki karar:
güncellemeyen kullanamasın).
"""
import threading
import tkinter as tk
from pathlib import Path

import ekran
import ikon


def read_version() -> str:
    try:
        return (Path(__file__).parent / "version.txt").read_text(encoding="utf-8").strip()
    except OSError:
        return "0.0.0"


def _dpi():
    """Tk()'den ÖNCE: Windows'ta bulanık yazıyı önler (Mac'te bir şey yapmaz)."""
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Yavuz.CrushCall")
    except Exception:
        pass


def _guncelleme_baslat() -> dict:
    kutu = {"bitti": False, "bilgi": None}

    def is_():
        try:
            import updater
            kutu["bilgi"] = updater.check_for_update(timeout=5)
        except Exception:
            pass
        kutu["bitti"] = True
    threading.Thread(target=is_, daemon=True).start()
    return kutu


def _zorunlu_guncelleme(kok, bilgi):
    from tkinter import messagebox
    import updater
    pen = tk.Toplevel(kok)
    pen.title("Updating CrushCall")
    pen.transient(kok)
    pen.grab_set()
    tk.Label(pen, text=f"Updating CrushCall to {bilgi.version}…\nIt opens again by itself.",
             padx=34, pady=21).pack()
    pen.update()
    try:
        ok = updater.run_update_flow(bilgi, parent_window=pen)
    except SystemExit:
        raise
    except Exception:
        ok = False
    if not ok:
        messagebox.showerror("Update required", "CrushCall must be updated to keep working, but the update "
                                                "could not be installed.\n\nCheck the internet connection and open CrushCall again.")
        kok.destroy()


def main():
    _dpi()
    kok = tk.Tk()
    ikon.set_icon(kok)
    ekran.Pencere(kok, surum=read_version())
    kutu = _guncelleme_baslat()     # iş parçacığı en geç ~5 sn'de biter (timeout)

    def kontrol():
        if not kutu["bitti"]:
            kok.after(200, kontrol)
        elif kutu["bilgi"]:
            _zorunlu_guncelleme(kok, kutu["bilgi"])
    kok.after(200, kontrol)
    kok.mainloop()


def selftest() -> int:
    """Pencere göstermeden bütün modülleri yükle, ekranı kur, şablonu aç. Hata = istisna."""
    import ayar, cikti_davet, cikti_not, islem, kaynak, pano, updater  # noqa: F401,E401
    from openpyxl import load_workbook
    load_workbook(cikti_davet.SABLON)
    kok = tk.Tk()
    kok.withdraw()
    try:
        ekran.Pencere(kok, surum=read_version())
        kok.update()
    finally:
        kok.destroy()
    return 0


if __name__ == "__main__":
    main()
