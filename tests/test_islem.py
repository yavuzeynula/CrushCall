from datetime import date

import islem
from tests.fixtures import BETON_BASLIK, D25, beton_satir, gun, kaynak_yap
from veri import Numune, Satir


def S(sid, ria="R", n=5, grup="concrete", yas="7d", el="Pile", sec="2", dokum=None, ch="1+000"):
    return Satir(grup=grup, yas=yas, sample_id=sid, test_tarihi=D25, kaynak_satir=n, ria=ria,
                 element=el, section=sec, chainage=ch, structure="Bridge", dokum=dokum)


def test_double_birlesir():
    ns, uy = islem.birlestir([S("G26-CON-1714", "R1", 5, ch="15+020"), S("G26-CON-1714", "R2", 6, ch="0+127")])
    assert len(ns) == 1 and uy == []
    n = ns[0]
    assert (n.ria, n.chainage, n.structure, n.set_sayisi) == ("R1\nR2", "15+020\n0+127", "Bridge", 2)


def test_ardisik_olmayan_double():
    ns, uy = islem.birlestir([S("A-1714", n=5), S("A-1715", n=6), S("A-1714", n=7)])
    assert [n.sample_id for n in ns] == ["A-1714", "A-1715"] and ns[0].set_sayisi == 2
    assert any("A-1714" in u.metin and "not next to each other" in u.metin for u in uy)


def test_section_birlesir():
    ns, _ = islem.birlestir([S("A-1", sec="2"), S("A-1", sec="1"), S("A-2", sec="2"), S("A-2", sec="2")])
    assert [n.section for n in ns] == ["2 & 1", "2"]


def test_shotcrete():
    ns, _ = islem.birlestir([S("A-1", el="Concrete for Slope 1"), S("A-2", el="SHOTCRETE for S8"), S("A-3", el="Pile")])
    assert [n.shotcrete for n in ns] == [True, True, False]


def test_no_ria():
    ns, _ = islem.birlestir([S("A-1", ria=""), S("A-2", ria="0"), S("A-3", ria=" 0 "), S("A-4", ria="R"),
                             S("C-1", ria="", grup="cmd"), S("T-1", ria="", grup="tcon")])
    assert [n.no_ria for n in ns] == [True, True, True, False, False, False]


def test_grout_yasi():
    n = lambda d: Numune("grout_s", "1or7", "G26-GRT-0254", [S("G26-GRT-0254", grup="grout_s", yas="1or7", dokum=d)])
    assert islem.grout_yasi(n(date(2026, 11, 24)), D25) == ("1d", None)
    assert islem.grout_yasi(n(date(2026, 11, 18)), D25) == ("7d", None)
    yas, u = islem.grout_yasi(n(date(2026, 11, 20)), D25)
    assert yas == "?" and "cannot tell 1 or 7 days" in u.metin
    assert islem.grout_yasi(n(None), D25)[0] == "?"


def _ns(*ids):
    return [Numune("concrete", "7d", i, []) for i in ids]


def test_seri_eksik():
    uy = islem.seri_kontrol("Concrete 7d", _ns("G26-CON-1713", "G26-CON-1715", "G26-CON-1718"), None)
    assert [u.metin for u in uy] == ["Concrete 7d: G26-CON-1714 missing", "Concrete 7d: G26-CON-1716 missing",
                                     "Concrete 7d: G26-CON-1717 missing"]


def test_seri_devam_ok():
    assert islem.seri_kontrol("Concrete 7d", _ns("G26-CON-1713", "G26-CON-1714"),
                              (date(2026, 11, 24), ["G26-CON-1711", "G26-CON-1712"])) == []


def test_seri_devam_kopuk():
    uy = islem.seri_kontrol("Concrete 7d", _ns("G26-CON-1715"), (date(2026, 11, 24), ["G26-CON-1712"]))
    assert [u.metin for u in uy] == [
        "Concrete 7d: previous day (24.11.2026) ended at G26-CON-1712, today starts at G26-CON-1715"]


def test_seri_buyuk_bosluk_tek_satir():
    uy = islem.seri_kontrol("Concrete 7d", _ns("G26-CON-1713", "G26-CON-2900"), None)
    assert len(uy) == 1 and "1186 numbers missing" in uy[0].metin


def test_seri_sayisiz_id_atlanir():
    assert islem.seri_kontrol("CMD 28d", _ns("CMD-A", "CMD-B"), None) == []


def test_hesapla_uctan_uca(tmp_path):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, [
        beton_satir("R", "G26-CON-1712", t7=gun(24)),
        beton_satir("R1", "G26-CON-1713", t7=gun(25)),
        beton_satir("R2", "G26-CON-1714", t7=gun(25), el="Shotcrete"),
        beton_satir("R3", "G26-CON-1714", t7=gun(25), el="Shotcrete"),
        beton_satir("", "G26-CON-1715", t7=gun(25)),
    ])
    s = islem.hesapla(D25, {"concrete": y}, gruplar=("concrete",))
    assert s.tamam, [u.metin for u in s.uyarilar]
    assert [(n.sample_id, n.set_sayisi, n.shotcrete, n.no_ria) for n in s.numuneler] == [
        ("G26-CON-1713", 1, False, False), ("G26-CON-1714", 2, True, False), ("G26-CON-1715", 1, False, True)]
    h = [h for h in s.hesap if h.ad == "Concrete 7d"][0]
    assert (h.kaynak_satir, h.numune, h.birlesen, h.no_ria) == (4, 3, 1, 1)
    assert not [u for u in s.uyarilar if "previous day" in u.metin]


def test_seri_kontrolu_yalniz_beton(tmp_path):
    # Grout numaraları yaşa göre sıralı değil (0256 dün, 0254 bugün: 2026-11-24 ekran) -> uyarı yok
    b = ["RIA No:", "Sample ID", "Pouring Date", "Location", None, None, None,
         "1-day or 7-days Compressive Strength EN 196-1", "28-days Compressive Strength EN 196-1"]
    y = kaynak_yap(tmp_path / "g.xlsx", b, [
        ["R", "G26-GRT-0256", gun(19), "2", "c", "s", "e", gun(20), None],
        ["R", "G26-GRT-0254", gun(18), "2", "c", "s", "e", gun(25), None],
    ], sayfa="Grout Comp Log")
    s = islem.hesapla(D25, {"grout_s": y}, gruplar=("grout_s",))
    assert [u.metin for u in s.uyarilar] == []


def test_dosya_secilmemis_hata():
    s = islem.hesapla(D25, {}, gruplar=("concrete",))
    assert not s.tamam and any("no source file chosen" in u.metin for u in s.uyarilar)
