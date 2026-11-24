from datetime import date

from openpyxl import load_workbook

import kaynak
from tanimlar import GRUP
from tests.fixtures import BETON_BASLIK, D25, beton_satir, gun, kaynak_yap


def _hatalar(o, seviye):
    return [u.metin for u in o.uyarilar if u.seviye == seviye]


def test_tarihe_esit_satirlar_gelir(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [
        beton_satir("R1", "G26-CON-1713", t7=gun(25), sec="2", ch="12+180", st="External Channel", el="Concrete"),
        beton_satir("R2", "G26-CON-1714", t7=gun(26)),
        beton_satir("R3", "G26-CON-1594", t28=gun(25)),
    ])
    o = kaynak.oku(y, GRUP["concrete"], D25)
    assert [(s.sample_id, s.yas) for s in o.satirlar] == [("G26-CON-1713", "7d"), ("G26-CON-1594", "28d")]
    s = o.satirlar[0]
    assert (s.ria, s.section, s.chainage, s.structure, s.element) == ("R1", "2", "12+180", "External Channel", "Concrete")
    assert s.kaynak_satir == 5 and s.test_tarihi == D25
    assert o.sayim == {"7d": 1, "28d": 1}
    assert _hatalar(o, "error") == []


def test_metin_tarih(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [
        beton_satir("R1", "A-1", t7="25/11/2026"),
        beton_satir("R2", "A-2", t7="25.11.2026"),
        beton_satir("R3", "A-3", t7="11/25/2026"),
    ])
    o = kaynak.oku(y, GRUP["concrete"], D25)
    assert [s.sample_id for s in o.satirlar] == ["A-1", "A-2"]
    uy = _hatalar(o, "warning")
    assert len(uy) == 1 and "row 7" in uy[0] and "11/25/2026" in uy[0]


def test_excel_seri_sayisi_tarih(tmp_path):
    # Power Query Date.From sayıyı da tarih sayıyordu; aynısı
    seri = (D25 - date(1899, 12, 30)).days
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [beton_satir("R1", "A-1", t7=seri)])
    assert [s.sample_id for s in kaynak.oku(y, GRUP["concrete"], D25).satirlar] == ["A-1"]


def test_baslik_kaymis_hata(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [beton_satir("R1", "A-1", t7=gun(25))], ust_bos=4)
    o = kaynak.oku(y, GRUP["concrete"], D25)
    assert o.satirlar == []
    assert any("header" in h and "not found" in h for h in _hatalar(o, "error"))


def test_eksik_tarih_sutunu_hata(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK[:-1], [beton_satir("R1", "A-1", t7=gun(25))[:-1]])
    o = kaynak.oku(y, GRUP["concrete"], D25)
    assert any("28 days Compressive Strength EN12390-3" in h for h in _hatalar(o, "error"))


def test_grout_e_28_iki_yazim(tmp_path):
    b = ["RIA No:", "Sample ID", "Mix No", "Production Location", "Pouring Date", "Sampling Location",
         "Location", None, None, None, "1 day Compressive Strength EN 196-1", "28 day Compressive Strength EN 196-1"]
    y = kaynak_yap(tmp_path / "g.xlsx", b, [["R", "G26-GRT-0240", "GRT-1", "Site", gun(28, 10), "Site",
                                            "2", "8+166", "B", "PT", None, gun(25)]], sayfa="Grout Comp Log")
    o = kaynak.oku(y, GRUP["grout_e"], D25)
    assert [(s.sample_id, s.yas) for s in o.satirlar] == [("G26-GRT-0240", "28d")]
    assert o.satirlar[0].dokum == date(2026, 10, 28)


def test_formul_degersiz(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [beton_satir("R1", "A-1", t7="=E5+7", dokum=gun(18))])
    o = kaynak.oku(y, GRUP["concrete"], D25)   # openpyxl hesaplamaz: değer yok
    assert any("formula" in h and "row 5" in h for h in _hatalar(o, "error"))


def test_dosya_yok_hata(tmp_path):
    o = kaynak.oku(tmp_path / "yok.xlsx", GRUP["concrete"], D25)
    assert o.satirlar == [] and any("file not found" in h for h in _hatalar(o, "error"))


def test_sayfa_yok_hata(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [], sayfa="Baska")
    assert any("sheet 'Concrete Comp Log' not found" in h for h in _hatalar(kaynak.oku(y, GRUP["concrete"], D25), "error"))


def test_coklu_satir_ve_bosluk_korunur(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK,
                   [beton_satir("R1\nR2", "A-1", t7=gun(25), st=" External Channel", sec=2.0)])
    s = kaynak.oku(y, GRUP["concrete"], D25).satirlar[0]
    assert s.ria == "R1\nR2" and s.structure == " External Channel" and s.section == "2"


def test_sample_id_bos_uyari(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [beton_satir("R1", None, t7=gun(25))])
    o = kaynak.oku(y, GRUP["concrete"], D25)
    assert o.satirlar == [] and any("Sample ID is empty" in h for h in _hatalar(o, "error"))


def test_onceki_gun(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [
        beton_satir("R", "G26-CON-1709", t7=gun(22)),
        beton_satir("R", "G26-CON-1711", t7=gun(24)),
        beton_satir("R", "G26-CON-1712", t7=gun(24)),
        beton_satir("R", "G26-CON-1713", t7=gun(25)),
    ])
    o = kaynak.oku(y, GRUP["concrete"], D25)
    assert o.onceki["7d"] == (date(2026, 11, 24), ["G26-CON-1711", "G26-CON-1712"])
    assert "28d" not in o.onceki


def test_kaynaga_yazmaz(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [beton_satir("R1", "A-1", t7=gun(25))])
    once = y.read_bytes()
    kaynak.oku(y, GRUP["concrete"], D25)
    assert y.read_bytes() == once
    load_workbook(y)  # bozulmadı
