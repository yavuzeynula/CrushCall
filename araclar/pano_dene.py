"""Copy rows uçtan uca denemesi (yalnız Windows + Excel): panoya kopyala, Excel'e yapıştır, geri oku.

python araclar/pano_dene.py      (Mac'ten: python3 ~/tools/yayin/vm.py araclar/pano_dene.py)
"""
import ctypes
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(KOK / "src"))

import cikti_davet  # noqa: E402
import pano  # noqa: E402
from veri import Numune  # noqa: E402


def panodaki_bicimler() -> list[str]:
    u32 = ctypes.windll.user32
    u32.GetClipboardFormatNameW.argtypes = [ctypes.c_uint, ctypes.c_wchar_p, ctypes.c_int]
    adlar, f = [], 0
    u32.OpenClipboard(None)
    try:
        while True:
            f = u32.EnumClipboardFormats(f)
            if not f:
                break
            tampon = ctypes.create_unicode_buffer(128)
            n = u32.GetClipboardFormatNameW(f, tampon, 128)
            adlar.append(tampon.value if n else str(f))
    finally:
        u32.CloseClipboard()
    return adlar


def main() -> int:
    ns = [Numune("concrete", "7d", "G26-CON-1714", [None], ria="R-1\nR-2", section="2", chainage="15+020\n0+127",
                 structure="Bridge B0149", element="Shotcrete"),
          Numune("grout_s", "28d", "G26-GRT-0240", [None], ria="R-3", section="2", chainage="8+166",
                 structure="Bridge B0080", element="Post-Tensioning")]
    satirlar = cikti_davet.satirlar(ns, date(2026, 11, 25))
    bicimli = pano.kopyala(cikti_davet.html(satirlar), cikti_davet.tsv(satirlar))
    print("kopyala() bicimli:", bicimli)
    print("panodaki bicimler:", panodaki_bicimler())
    ps = r"""
$x = New-Object -ComObject Excel.Application; $x.DisplayAlerts = $false
$wb = $x.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
$ws.Range("A1").Select() | Out-Null; $ws.Paste()
foreach ($r in 1..3) { $s = @(); foreach ($c in 1..11) { $s += ($ws.Cells.Item($r,$c).Text -replace "`n","\n") }; "ROW $r : " + ($s -join " | ") }
"RENK D1 = " + $ws.Range("D1").Interior.Color + " ; RENK C2 = " + $ws.Range("C2").Interior.Color
$wb.Close($false); $x.Quit()
"""
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as f:
        f.write(ps)
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", f.name],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout, r.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
