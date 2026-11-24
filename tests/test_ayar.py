from pathlib import Path

import ayar


def test_yaz_oku(tmp_path, monkeypatch):
    monkeypatch.setattr(Path, "home", lambda: tmp_path)
    assert ayar.oku() == {}
    ayar.yaz({"kaynaklar": {"concrete": "C:\\x\\Beton ş.xlsx"}, "tema": "dark"})
    assert ayar.oku()["kaynaklar"]["concrete"] == "C:\\x\\Beton ş.xlsx"


def test_bozuk_dosya(tmp_path, monkeypatch):
    monkeypatch.setattr(Path, "home", lambda: tmp_path)
    d = ayar.dosya()
    d.parent.mkdir(parents=True)
    d.write_text("{bozuk", encoding="utf-8")
    assert ayar.oku() == {}
    assert d.with_name("ayarlar.json.bozuk").exists()
