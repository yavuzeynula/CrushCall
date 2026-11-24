"""Kabul testi: CrushCall'un çıktısını Yavuz'un elle hazırladığı geçmiş günlerle karşılaştırır.

python araclar/karsilastir.py --kaynaklar <5 log'un klasörü> --sablon "Daily Test Invitation TEMPLATE.xlsx"
                              --not <not dosyası.txt> [--gun 20] [--kayma 0]

- Kaynak dosyalar adlarından tanınır (Without RIA / DESIGNS / Elastomeric / Slope / diğer = Concrete).
- Şablondaki "(gg.aa.yyyy)" sayfaları: F = Sample ID, E = RIA, K–N = Section/Chainage/Structure/Element.
- Not dosyası: "///7days-gg.aa.yyyy\\\\\\" blokları ve altındaki "•" satırları.
- --kayma: sayfa adındaki tarih ile test tarihi arasındaki gün farkı (sayfa adı test gününü göstermiyorsa).
Her fark tek satır yazılır. Fark yoksa "0 differences", çıkış kodu 0.
"""
import argparse
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(KOK / "src"))

from openpyxl import load_workbook  # noqa: E402

import cikti_not  # noqa: E402
import islem  # noqa: E402
from kaynak import metin  # noqa: E402

_SAYFA = re.compile(r"^\((\d{2})\.(\d{2})\.(\d{4})\)$")
_BASLIK = re.compile(r"^///\d+days-(\d{2})\.(\d{2})\.(\d{4})\\\\\\$")
_ID = re.compile(r"^(?:[A-Z0-9]+-)?[A-Z]+-\d+$")


def kaynaklari_tani(klasor: Path) -> dict[str, Path]:
    secim = {}
    for y in sorted(Path(klasor).glob("*.xls*")):
        ad = y.name.lower()
        kod = ("tcon" if "without ria" in ad else "cmd" if "designs" in ad else
               "grout_e" if "elastomeric" in ad else "grout_s" if "slope" in ad else
               "concrete" if "concrete" in ad else None)
        if kod:
            secim[kod] = y
    return secim


def beklenen_davet(ws) -> dict[str, tuple[str, ...]]:
    """Sayfadaki her Sample ID için (RIA, Section, Chainage, Structure, Element)."""
    out = {}
    for satir in ws.iter_rows(min_row=1, values_only=True):
        sid = metin(satir[5] if len(satir) > 5 else None).strip()
        if _ID.match(sid):
            al = lambda i: metin(satir[i] if len(satir) > i else None).replace("\r\n", "\n")
            out[sid] = (al(4), al(10), al(11), al(12), al(13))
    return out


def beklenen_not(yazi: str) -> dict[date, list[str]]:
    """Not dosyasını güne göre böler (madde satırları son başlığın gününe aittir)."""
    out, gun = {}, None
    for satir in yazi.splitlines():
        s = satir.strip()
        m = _BASLIK.match(s)
        if m:
            gun = date(int(m[3]), int(m[2]), int(m[1]))
        if gun and s:
            out.setdefault(gun, []).append(s)
    return out


def gun_farklari(tarih: date, sonuc, davet: dict | None, notlar: list[str] | None) -> list[str]:
    f = []
    gun = f"{tarih:%d.%m.%Y}"
    for u in sonuc.uyarilar:
        if u.seviye == "error":
            f.append(f"{gun} program error: {u.metin}")
    if notlar is not None:
        bizim = [s.strip() for s in cikti_not.yap(sonuc.numuneler, tarih).splitlines()]
        if bizim != notlar:
            f.append(f"{gun} note differs:\n    expected: {notlar}\n    program:  {bizim}")
    if davet is not None:
        bizim = {n.sample_id: (n.ria, n.section, n.chainage, n.structure, n.element)
                 for n in sonuc.numuneler if n.grup in ("concrete", "grout_e", "grout_s")}
        for sid in sorted(set(davet) | set(bizim)):
            if sid not in bizim:
                f.append(f"{gun} {sid}: in your sheet, not found by program")
            elif sid not in davet:
                f.append(f"{gun} {sid}: found by program, not in your sheet")
            else:
                for ad, b, p in zip(("RIA", "Section", "Chainage", "Structure", "Element"), davet[sid], bizim[sid]):
                    if b.strip() != p.strip():
                        f.append(f"{gun} {sid} {ad}: expected {b!r}, program {p!r}")
    return f


def calistir(kaynaklar: Path, sablon: Path | None, not_dosyasi: Path | None, gun: int, kayma: int) -> list[str]:
    secim = kaynaklari_tani(kaynaklar)
    sayfalar = {}
    if sablon:
        wb = load_workbook(sablon, read_only=True, data_only=True)
        for ad in wb.sheetnames:
            m = _SAYFA.match(ad)
            if m:
                sayfalar[date(int(m[3]), int(m[2]), int(m[1])) + timedelta(days=kayma)] = beklenen_davet(wb[ad])
    notlar = beklenen_not(Path(not_dosyasi).read_text(encoding="utf-8", errors="replace")) if not_dosyasi else {}
    gunler = sorted(set(sayfalar) | set(notlar))[-gun:]
    farklar = []
    for t in gunler:
        farklar += gun_farklari(t, islem.hesapla(t, secim), sayfalar.get(t), notlar.get(t))
    return farklar


def main(argv=None) -> int:
    a = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    a.add_argument("--kaynaklar", type=Path, required=True)
    a.add_argument("--sablon", type=Path)
    a.add_argument("--not", dest="not_dosyasi", type=Path)
    a.add_argument("--gun", type=int, default=20)
    a.add_argument("--kayma", type=int, default=0)
    g = a.parse_args(argv)
    farklar = calistir(g.kaynaklar, g.sablon, g.not_dosyasi, g.gun, g.kayma)
    for f in farklar:
        print(f)
    print(f"{len(farklar)} differences")
    return 0 if not farklar else 1


if __name__ == "__main__":
    sys.exit(main())
