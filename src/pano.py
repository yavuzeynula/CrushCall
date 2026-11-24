"""Panoya biçimli tablo (HTML) + düz metin koyar; Excel HTML'i biçimiyle yapıştırır.

Windows: ctypes ile "HTML Format" + CF_UNICODETEXT (ek paket yok).
Mac: osascript.  Olmazsa: Tk panosuna yalnız düz metin.
"""
import subprocess
import sys

_BAS = ("Version:0.9\r\nStartHTML:{0:010d}\r\nEndHTML:{1:010d}\r\n"
        "StartFragment:{2:010d}\r\nEndFragment:{3:010d}\r\n")


def cf_html(parca: str) -> bytes:
    """Windows "HTML Format" verisi: başlıktaki sayılar bayt konumlarıdır."""
    on = "<html><body><!--StartFragment-->".encode("utf-8")
    govde = parca.encode("utf-8")
    son = "<!--EndFragment--></body></html>".encode("utf-8")
    bas_uzunluk = len(_BAS.format(0, 0, 0, 0).encode("ascii"))
    bas_html = bas_uzunluk
    bas_parca = bas_html + len(on)
    son_parca = bas_parca + len(govde)
    son_html = son_parca + len(son)
    return _BAS.format(bas_html, son_html, bas_parca, son_parca).encode("ascii") + on + govde + son


def _windows(html: str, metin: str) -> None:
    import ctypes
    from ctypes import wintypes
    u32, k32 = ctypes.windll.user32, ctypes.windll.kernel32
    u32.OpenClipboard.argtypes = [wintypes.HWND]
    u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    u32.SetClipboardData.restype = wintypes.HANDLE
    u32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    u32.RegisterClipboardFormatW.restype = wintypes.UINT
    k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k32.GlobalAlloc.restype = wintypes.HGLOBAL
    k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalLock.restype = wintypes.LPVOID
    k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]

    def koy(bicim: int, veri: bytes) -> None:
        h = k32.GlobalAlloc(0x0002, len(veri))           # GMEM_MOVEABLE
        p = k32.GlobalLock(h)
        ctypes.memmove(p, veri, len(veri))
        k32.GlobalUnlock(h)
        if not u32.SetClipboardData(bicim, h):
            raise OSError("SetClipboardData failed")

    if not u32.OpenClipboard(None):
        raise OSError("clipboard is busy")
    try:
        u32.EmptyClipboard()
        koy(u32.RegisterClipboardFormatW("HTML Format"), cf_html(html) + b"\0")
        koy(13, (metin + "\0").encode("utf-16-le"))      # CF_UNICODETEXT
    finally:
        u32.CloseClipboard()


def _mac(html: str, metin: str) -> None:
    betik = ("set the clipboard to {«class HTML»:«data HTML" + html.encode("utf-8").hex() + "», "
             "«class utf8»:«data utf8" + metin.encode("utf-8").hex() + "»}")
    subprocess.run(["osascript", "-e", betik], check=True, capture_output=True)


def kopyala(html: str, metin: str, tk_kok=None) -> bool:
    """True: biçimli kopyalandı. False: yalnız düz metin (Tk panosu)."""
    try:
        if sys.platform == "win32":
            _windows(html, metin)
            return True
        if sys.platform == "darwin":
            _mac(html, metin)
            return True
    except Exception:
        pass
    if tk_kok is not None:
        tk_kok.clipboard_clear()
        tk_kok.clipboard_append(metin)
    return False


# --- Excel'in kendi kopyası (Windows) ---------------------------------------
# Excel "Insert Copied Cells" seçeneğini yalnız kendi içinden yapılan kopyada
# gösterir; dışarıdan gelen HTML'de yalnız "Paste" çıkar (Yavuz 2026-11-24).
# Bu yüzden dosya kullanıcının açık Excel'inde açılır ve satırları Excel kopyalar.
_EXCEL_BETIK = r"""
$ErrorActionPreference = 'Stop'
$yol = $env:CRUSHCALL_DOSYA; $aralik = $env:CRUSHCALL_ARALIK
try { $x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { $x = New-Object -ComObject Excel.Application }
$x.Visible = $true
$wb = $null
foreach ($w in $x.Workbooks) { if ($w.FullName -eq $yol) { $wb = $w } }
if (-not $wb) { $wb = $x.Workbooks.Open($yol) }
$ws = $wb.Worksheets.Item(1)
$ws.Activate() | Out-Null
$r = $ws.Range($aralik)
$r.Select() | Out-Null
$r.Copy() | Out-Null
'OK'
"""


def excel_aralik(satir_sayisi: int, ilk: int = 13) -> str:
    """Davet dosyasında sarı veri satırları (C–M sütunları)."""
    return f"C{ilk}:M{ilk + satir_sayisi - 1}"


def excel_kopyala(dosya, satir_sayisi: int) -> bool:
    """Dosyayı açık Excel'de açıp sarı satırları Excel'e kopyalatır. Yalnız Windows; olmazsa False."""
    if sys.platform != "win32" or satir_sayisi < 1:
        return False
    import os
    ortam = {**os.environ, "CRUSHCALL_DOSYA": str(dosya), "CRUSHCALL_ARALIK": excel_aralik(satir_sayisi)}
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", _EXCEL_BETIK],
                           env=ortam, capture_output=True, text=True, timeout=90,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired):
        return False
    return r.returncode == 0 and r.stdout.strip().endswith("OK")
