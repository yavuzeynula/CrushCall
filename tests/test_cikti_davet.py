from datetime import time
from pathlib import Path

import pytest
from openpyxl import load_workbook

import cikti_davet
import pano
from tests.fixtures import D25
from veri import Numune

ORNEK = Path(__file__).resolve().parent.parent / "ornekler" / "Laboratory_Daily_Notification_Invitation_List_25.11.2026.xlsx"


def _ornek_numuneler():
    """Örnek dosyadaki 16 satırı Numune'ye geri çevirir (yalnız test için)."""
    ws = load_workbook(ORNEK).active
    ns = []
    for r in range(13, 29):
        v = [ws.cell(r, c).value for c in range(3, 14)]
        test = v[2]
        grup = "grout_s" if test.startswith("Grout") else "concrete"
        yas = "7d" if "(7d)" in test else "28d"
        t = lambda x: (x or "").replace("\r\n", "\n")
        ns.append(Numune(grup, yas, f"ID-{r}", [None], ria=t(v[4]), section=t(str(v[6])), chainage=t(v[7]),
                         structure=t(v[8]), element=t(v[9])))
    return ns, ws


@pytest.mark.skipif(not ORNEK.exists(), reason="ornekler/ yok (gerçek veri depoda değil)")
def test_satirlar_ornekle_ayni():
    ns, ws = _ornek_numuneler()
    satirlar = cikti_davet.satirlar(ns, D25)
    assert len(satirlar) == 16
    for i, s in enumerate(satirlar):
        r = 13 + i
        beklenen = [ws.cell(r, c).value for c in range(3, 14)]
        beklenen = [(b.replace("\r\n", "\n") if isinstance(b, str) else b) for b in beklenen]
        beklenen[0] = int(beklenen[0])
        beklenen[3] = "25.11.2026"                         # örnekte formül, sonucu bu
        beklenen[6] = str(beklenen[6])
        beklenen[10] = "09:00"                             # örnekte Excel saati 09:00
        assert s == beklenen, f"satır {r}"


def test_no_ria_ve_tcon_girmez():
    ns = [Numune("concrete", "7d", "A-1", [None], ria=""), Numune("tcon", "7d", "T-1", [None]),
          Numune("concrete", "7d", "A-2", [None], ria="R")]
    ns[0].no_ria = True
    assert [s[4] for s in cikti_davet.satirlar(ns, D25)] == ["R"]


def test_cmd_satiri():
    n = Numune("cmd", "28d", "CMD-0123", [None], mix_no="CMD-8011", sinif="C35/45")
    assert cikti_davet.satirlar([n], D25) == [[1, "BEJV LAB", "Compressive Strength (28d) ", "25.11.2026", "-", "-",
                                               "1", "11+500\nQC Lab", "-", "Concrete Mix design  #CMD-8011 C35/45",
                                               "11:00"]]


def test_sira_beton_cmd_grout():
    ns = [Numune("grout_e", "28d", "G", [None], ria="R"), Numune("cmd", "7d", "C", [None]),
          Numune("concrete", "28d", "B28", [None], ria="R"), Numune("concrete", "7d", "B7", [None], ria="R")]
    assert [s[2].split("(")[1] for s in cikti_davet.satirlar(ns, D25)] == ["7d) ", "28d) ", "7d) ", "28d) "]
    assert cikti_davet.satirlar(ns, D25)[3][2] == "Grout \nCompressive Strength (28d) "


def test_kaydet_bicim(tmp_path):
    ns = [Numune("concrete", "7d", "A", [None], ria="R1\nR2\nR3", element="E"),
          Numune("grout_s", "28d", "G", [None], ria="R")]
    yol = cikti_davet.kaydet(cikti_davet.satirlar(ns, D25), D25, tmp_path)
    assert yol.name == "Laboratory_Daily_Notification_Invitation_List_25.11.2026.xlsx"
    wb = load_workbook(yol)
    ws = wb.active
    assert ws.title == "25.11.2026" and len(ws._images) == 3
    assert ws["D13"].fill.fgColor.rgb == "FFFFFF00" and ws["C13"].fill.fgColor.rgb != "FFFFFF00"
    assert ws["E14"].fill.fgColor.rgb == "FF92D050" and ws["E13"].fill.fgColor.rgb == "FFFFFF00"
    assert ws["G13"].border.left.style == "thin" and ws["G13"].alignment.wrap_text
    assert ws["M13"].value == time(9, 0) and ws["M13"].number_format == "h:mm"
    assert ws.row_dimensions[13].height == 45 and ws.row_dimensions[14].height == 30
    assert ws.max_row == 14          # önce: hücreye bakmak hücreyi oluşturur


def test_kaydet_dosya_aciksa_anlasilir_hata(tmp_path, monkeypatch):
    def kilitli(*a, **k):
        raise PermissionError("locked")
    monkeypatch.setattr("openpyxl.workbook.workbook.Workbook.save", kilitli)
    with pytest.raises(cikti_davet.KayitHatasi, match="close it in Excel"):
        cikti_davet.kaydet([], D25, tmp_path)


def test_html_ve_tsv_baslik_yok():
    satirlar = cikti_davet.satirlar([Numune("grout_e", "7d", "G", [None], ria="R1\nR2")], D25)
    h = cikti_davet.html(satirlar)
    assert "Laboratory" not in h and h.count("<tr") == 1
    assert 'R1<br style="mso-data-placement:same-cell">R2' in h
    assert "#92D050" in h and "#FFFF00" in h
    t = cikti_davet.tsv(satirlar)
    assert t.count("\r\n") == 0 and '"R1\nR2"' in t and '"Grout \nCompressive Strength (7d) "' in t


def test_cf_html_ofsetleri():
    veri = pano.cf_html("<table><tr><td>x</td></tr></table>")
    metin = veri.decode("utf-8")
    al = lambda ad: int(metin.split(ad + ":")[1][:10])
    assert veri[al("StartFragment"):al("EndFragment")].decode() == "<table><tr><td>x</td></tr></table>"
    assert veri[al("StartHTML"):al("EndHTML")].decode().startswith("<html>")
