"""Copy rows (Excel'in kendi kopyası) uçtan uca denemesi — yalnız Windows + Excel.

python3 ~/tools/yayin/vm.py araclar/excel_kopya_dene.py
Yavuz'un durumu: kendi davet dosyası (aynı adla, başka klasörde) Excel'de açık.
Copy rows -> o dosyada 13. satır başlığına "Insert Copied Cells" (Rows.Insert) -> değer + satır yüksekliği okunur.
"""
import shutil
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

TARIH = date(2026, 11, 25)

KULLANICI_AC = r"""
try { $x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { $x = New-Object -ComObject Excel.Application }
$x.Visible = $true
$x.Workbooks.Open($env:KULLANICI) | Out-Null
'acik: ' + ($x.Workbooks | ForEach-Object { $_.Name }) -join ', '
"""

DENETIM = r"""
$x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
'acik kitaplar: ' + (($x.Workbooks | ForEach-Object { $_.Name }) -join ', ')
"CutCopyMode (1 = copy hazir) = " + $x.CutCopyMode
$hedef = $null; foreach ($w in $x.Workbooks) { if ($w.FullName -eq $env:KULLANICI) { $hedef = $w } }
$ws = $hedef.Worksheets.Item(1)
$ws.Rows("13:13").Insert(-4121) | Out-Null        # satir basligi > Insert Copied Cells
foreach ($r in 12..16) { $s = @(); foreach ($c in 3..13) { $s += ($ws.Cells.Item($r,$c).Text -replace "`n","\n") }; "ROW $r (yukseklik " + $ws.Rows($r).RowHeight + ") : " + ($s -join " | ") }
$x.CutCopyMode = $false
foreach ($w in @($x.Workbooks)) { $w.Close($false) }
$x.Quit()
"""


def ps(betik: str, **ortam) -> None:
    import os
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", betik],
                       env={**os.environ, **ortam}, capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout, r.stderr)


def main() -> int:
    ns = [Numune("concrete", "7d", "G26-CON-1714", [None], ria="R-1\nR-2", section="2", chainage="15+020\n0+127",
                 structure="Bridge B0149", element="Shotcrete"),
          Numune("grout_s", "28d", "G26-GRT-0240", [None], ria="R-3", section="2", chainage="8+166",
                 structure="Bridge B0080", element="Post-Tensioning")]
    satirlar = cikti_davet.satirlar(ns, TARIH)
    # Yavuz'un kendi dosyası: aynı ad, başka klasör, şablondan (boş).
    kullanici = Path(tempfile.mkdtemp()) / cikti_davet.dosya_adi(TARIH)
    shutil.copy(cikti_davet.SABLON, kullanici)
    ps(KULLANICI_AC, KULLANICI=str(kullanici))
    print("excel_kopyala:", pano.excel_kopyala(cikti_davet.kaydet(satirlar, TARIH, Path(tempfile.mkdtemp())), len(satirlar)))
    ps(DENETIM, KULLANICI=str(kullanici))
    return 0


if __name__ == "__main__":
    sys.exit(main())
