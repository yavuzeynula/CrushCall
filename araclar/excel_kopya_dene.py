"""Copy rows (Excel'in kendi kopyası) uçtan uca denemesi — yalnız Windows + Excel.

python3 ~/tools/yayin/vm.py araclar/excel_kopya_dene.py
Kaydet -> pano.excel_kopyala -> aynı Excel'de ikinci dosyada "Insert Copied Cells" (Range.Insert) -> geri oku.
"""
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

DENETIM = r"""
$x = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
"CutCopyMode (1 = copy hazir) = " + $x.CutCopyMode
$hedef = $x.Workbooks.Add(); $ws = $hedef.Worksheets.Item(1)
$ws.Range("A1").Value2 = "eski 1"; $ws.Range("A2").Value2 = "eski 2"
$ws.Range("A2").Insert(-4121) | Out-Null          # xlShiftDown = Insert Copied Cells
foreach ($r in 1..4) { $s = @(); foreach ($c in 1..11) { $s += ($ws.Cells.Item($r,$c).Text -replace "`n","\n") }; "ROW $r : " + ($s -join " | ") }
"RENK B2 = " + $ws.Range("B2").Interior.Color + " ; RENK C3 = " + $ws.Range("C3").Interior.Color
$x.CutCopyMode = $false
$hedef.Close($false)
foreach ($w in @($x.Workbooks)) { if ($w.Name -like "Laboratory_Daily*") { $w.Close($false) } }
if ($x.Workbooks.Count -eq 0) { $x.Quit() }
"""


def main() -> int:
    ns = [Numune("concrete", "7d", "G26-CON-1714", [None], ria="R-1\nR-2", section="2", chainage="15+020\n0+127",
                 structure="Bridge B0149", element="Shotcrete"),
          Numune("grout_s", "28d", "G26-GRT-0240", [None], ria="R-3", section="2", chainage="8+166",
                 structure="Bridge B0080", element="Post-Tensioning")]
    satirlar = cikti_davet.satirlar(ns, date(2026, 11, 25))
    yol = cikti_davet.kaydet(satirlar, date(2026, 11, 25), Path(tempfile.mkdtemp()))
    print("excel_kopyala:", pano.excel_kopyala(yol, len(satirlar)))
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", DENETIM],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(r.stdout, r.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
