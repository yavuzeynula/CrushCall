import sys
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "araclar"))
import karsilastir  # noqa: E402
from tests.fixtures import BETON_BASLIK, beton_satir, gun, kaynak_yap  # noqa: E402


def _hazirla(tmp_path, sablon_satirlari, not_yazisi):
    k = tmp_path / "kaynak"
    k.mkdir()
    kaynak_yap(k / "SECTION 1-2 Concrete Compression Strength.xlsx", BETON_BASLIK, [
        beton_satir("R", "G26-CON-1712", t7=gun(24)),
        beton_satir("R1", "G26-CON-1713", t7=gun(25), sec="2", ch="1+000", st="Bridge", el="Pile"),
    ])
    wb = Workbook()
    ws = wb.active
    ws.title = "(25.11.2026)"
    for s in sablon_satirlari:
        ws.append(s)
    wb.save(tmp_path / "sablon.xlsx")
    (tmp_path / "not.txt").write_text(not_yazisi, encoding="utf-8")
    return k


def _farklar(tmp_path, k):
    fs = karsilastir.calistir(k, tmp_path / "sablon.xlsx", tmp_path / "not.txt", 20, 0)
    return [f for f in fs if "no source file chosen" not in f]   # testte yalnız beton dosyası var


def test_fark_yok(tmp_path):
    k = _hazirla(tmp_path, [[None, None, None, None, "R1", "G26-CON-1713", None, None, None, None, "2", "1+000", "Bridge", "Pile"]],
                 "///7days-25.11.2026\\\\\\\n1. G26-CON-1713\n")
    assert _farklar(tmp_path, k) == []


def test_fark_raporu(tmp_path):
    k = _hazirla(tmp_path, [[None, None, None, None, "R9", "G26-CON-1713", None, None, None, None, "2", "1+000", "Bridge", "Pile"],
                            [None, None, None, None, "R", "G26-CON-1799"]],
                 "///7days-25.11.2026\\\\\\\n1. G26-CON-1713 SHOTCRETE\n")
    fs = _farklar(tmp_path, k)
    assert any("G26-CON-1713 RIA: expected 'R9', program 'R1'" in f for f in fs)
    assert any("G26-CON-1799: in your sheet, not found by program" in f for f in fs)
    assert any("note differs" in f for f in fs)


def test_kaynaklari_tani(tmp_path):
    for ad in ("Concrete Compression Strength (conformity) (implemented) - Without RIA.xlsx",
               "Concrete Compression Strength (conformity) (DESIGNS).xlsx",
               "SECTION 1-2 Grout Compression Strength for Elastomeric Bearing and Pillows.xlsx",
               "SECTION 1-2 Grout Compression Strength for Slope Protection Anchors and Post-tensioning.xlsx",
               "SECTION 1-2 Concrete Compression Strength (conformity) (implemented) (updated) from Sheet 04050 -.xlsx"):
        (tmp_path / ad).write_bytes(b"")
    assert set(karsilastir.kaynaklari_tani(tmp_path)) == {"tcon", "cmd", "grout_e", "grout_s", "concrete"}
