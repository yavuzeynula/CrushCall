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
# Dosya çift tıklar gibi (kabukla) açılır: COM ile "açık Excel"i aramak, kullanıcı Excel'i
# henüz kayıt etmemişse ikinci bir Excel başlatıyordu; kopya oradaydı, kullanıcının
# dosyasında yalnız "Insert" (boş satır) çıkıyordu (Yavuz 2026-11-25, VM'de doğrulandı).
# Kitap, Windows'un açık nesne tablosundan dosya yoluyla alınır (Excel başlatmaz).
# Tüm satır kopyalanır: satır başlığına yapıştırılabilir, satır yüksekliği de gelir.
# Excel aynı adlı iki kitabı açamaz: önceki CrushCall kopyası hâlâ açıksa kapatılır;
# önce kopya bırakılır, yoksa Excel "panoda çok veri var" sorusunda takılır.
_EXCEL_BETIK = r"""
$ErrorActionPreference = 'Stop'
$yol = $env:CRUSHCALL_DOSYA; $aralik = $env:CRUSHCALL_ARALIK; $onceki = $env:CRUSHCALL_ONCEKI
# Açık kitabı Windows'un açık nesne tablosundan alır; açık değilse null (BindToMoniker gibi Excel başlatmaz).
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class CrushRot {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int r, out IRunningObjectTable t);
  [DllImport("ole32.dll", CharSet = CharSet.Unicode)] static extern int CreateFileMoniker(string yol, out IMoniker m);
  public static object Al(string yol) {
    IRunningObjectTable t; IMoniker m; object o;
    if (GetRunningObjectTable(0, out t) != 0 || CreateFileMoniker(yol, out m) != 0) return null;
    return t.GetObject(m, out o) == 0 ? o : null;
  }
}
'@
function Dene($is) {
  for ($i = 0; $i -lt 50; $i++) { try { return & $is } catch { Start-Sleep -Milliseconds 200 } }
  & $is
}
if ($onceki) {
  $o = [CrushRot]::Al($onceki)
  if ($o) { Dene { $o.Application.CutCopyMode = $false; $o.Close($false) } | Out-Null }
}
Invoke-Item -LiteralPath $yol              # çift tıklar gibi: kullanıcının açık Excel'ine gider
$wb = $null
for ($i = 0; $i -lt 600 -and -not $wb; $i++) { Start-Sleep -Milliseconds 100; $wb = [CrushRot]::Al($yol) }
if (-not $wb) { throw 'Excel workbook not found' }
Dene { $ws = $wb.Worksheets.Item(1); $ws.Activate() | Out-Null; $r = $ws.Rows($aralik); $r.Select() | Out-Null; $r.Copy() | Out-Null } | Out-Null
'OK'
"""
_onceki = None   # bu oturumda son açılan kopya dosyası


def excel_aralik(satir_sayisi: int, ilk: int = 13) -> str:
    """Davet dosyasında veri satırları (tüm satır)."""
    return f"{ilk}:{ilk + satir_sayisi - 1}"


def excel_kopyala(dosya, satir_sayisi: int) -> bool:
    """Dosyayı açık Excel'de açıp veri satırlarını Excel'e kopyalatır. Yalnız Windows; olmazsa False."""
    if sys.platform != "win32" or satir_sayisi < 1:
        return False
    import os
    global _onceki
    ortam = {**os.environ, "CRUSHCALL_DOSYA": str(dosya), "CRUSHCALL_ARALIK": excel_aralik(satir_sayisi),
             "CRUSHCALL_ONCEKI": str(_onceki or "")}
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", _EXCEL_BETIK],
                           env=ortam, capture_output=True, text=True, timeout=90,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired):
        return False
    if r.returncode == 0 and r.stdout.strip().endswith("OK"):
        _onceki = dosya
        return True
    return False
