import subprocess
import sys
from pathlib import Path

import ayar
import ekran
import islem
from tests.fixtures import BETON_BASLIK, D25, beton_satir, gun, kaynak_yap

KOK = Path(__file__).resolve().parent.parent


def _pencere(tk_root, tmp_path, satirlar, gruplar=None):
    y = kaynak_yap(tmp_path / "b.xlsx", BETON_BASLIK, satirlar)
    ayar.yaz({"kaynaklar": {"concrete": str(y)}})
    p = ekran.Pencere(tk_root)
    p.tarih = D25
    p.tarih_var.set("25.11.2026")
    # Pull'un iş parçacığı yerine aynı hesap doğrudan (olay döngüsü beklenmez)
    p.sonuc = islem.hesapla(D25, ayar.oku()["kaynaklar"], gruplar=gruplar)
    p._goster()
    return p


def test_selftest_ayri_surecte_0():
    # yayındaki gibi: ayrı süreçte, pencere göstermeden
    kod = "import sys; sys.path.insert(0, 'src'); import main; sys.exit(main.selftest())"
    r = subprocess.run([sys.executable, "-c", kod], cwd=KOK, capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr


def test_pull_sonuc_gosterir(tk_root, tmp_path):
    p = _pencere(tk_root, tmp_path, [beton_satir("R1", "G26-CON-1713", t7=gun(25)),
                                     beton_satir("R2", "G26-CON-1714", t7=gun(25))], gruplar=("concrete",))
    metin = p.kontrol.get("1.0", "end")
    assert "Concrete 7d · 2" in metin and "2 source rows → 2 samples" in metin
    assert "1. G26-CON-1713" in p.not_metni.get("1.0", "end")
    p.sekme_sec("invite", gecis=False)
    assert str(p.b_kaydet.cget("state")) == "normal"
    assert len(p.tablo.find_all()) > 0


def test_hata_varken_kaydet_kapali(tk_root, tmp_path):
    p = _pencere(tk_root, tmp_path, [beton_satir("R1", "G26-CON-1713", t7=gun(25))])  # diğer gruplar seçilmemiş
    p.sekme_sec("invite", gecis=False)
    assert not p.sonuc.tamam
    assert str(p.b_kaydet.cget("state")) == "disabled" and str(p.b_satir.cget("state")) == "disabled"
    assert "no source file chosen" in p.kontrol.get("1.0", "end")


def test_tarih_okunur(tk_root, tmp_path):
    p = _pencere(tk_root, tmp_path, [], gruplar=("concrete",))
    p.tarih_var.set("01.12.2026")
    p.tarih_yazildi()
    assert f"{p.tarih:%d.%m.%Y}" == "01.12.2026"
    p.tarih_var.set("2026-12-01")
    p.tarih_yazildi()
    assert f"{p.tarih:%d.%m.%Y}" == "01.12.2026" and "dd.mm.yyyy" in p.durum.cget("text")


def test_buyut(tk_root, monkeypatch):
    import tkinter as tk
    import main
    cagri = []
    monkeypatch.setattr(tk_root, "state", lambda *a: cagri.append(a))
    main.buyut(tk_root)
    assert cagri == [("zoomed",)]

    def desteklemez(*a):
        raise tk.TclError("bad argument")
    monkeypatch.setattr(tk_root, "state", desteklemez)
    olcu = []
    monkeypatch.setattr(tk_root, "geometry", lambda g: olcu.append(g))
    main.buyut(tk_root)
    assert olcu == [f"{tk_root.winfo_screenwidth()}x{tk_root.winfo_screenheight()}+0+0"]
