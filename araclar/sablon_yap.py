"""Davet şablonunu örnek dosyadan üretir (bir kez; örnek değişirse yeniden).

python araclar/sablon_yap.py <örnek davet listesi .xlsx>
Üst kısım ve 13. satırın biçimi kalır; veri satırları silinir.
"""
import sys
from pathlib import Path

from openpyxl import load_workbook

KOK = Path(__file__).resolve().parent.parent
HEDEF = KOK / "src" / "sablon" / "davet_sablonu.xlsx"


def yap(ornek: Path, hedef: Path = HEDEF) -> Path:
    wb = load_workbook(ornek)
    ws = wb.active
    for c in range(3, 14):
        ws.cell(13, c).value = None
    ws.delete_rows(14, ws.max_row)
    ws.title = "TEMPLATE"
    hedef.parent.mkdir(parents=True, exist_ok=True)
    wb.save(hedef)
    return hedef


if __name__ == "__main__":
    print(yap(Path(sys.argv[1])))
