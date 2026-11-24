"""Okunan satırları numuneye çevirir: birleştirme, işaretler, grout yaşı, seri kontrolü, hesap."""
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import kaynak
from tanimlar import GRUP, GRUPLAR
from veri import Numune, Satir, Uyari

_SAYI = re.compile(r"^(.*?)(\d+)$")
EN_COK_EKSIK = 10      # bundan çok eksik numara varsa tek satırda söylenir


@dataclass
class Hesap:
    """Bir grup + yaş için her satırın nereye gittiği."""
    ad: str
    kaynak_satir: int      # kaynakta tarihi eşleşen hücre sayısı
    numune: int
    birlesen: int          # birleştirilen fazla satır
    no_ria: int
    sadece_not: bool


@dataclass
class Sonuc:
    numuneler: list[Numune] = field(default_factory=list)
    uyarilar: list[Uyari] = field(default_factory=list)
    hesap: list[Hesap] = field(default_factory=list)

    @property
    def tamam(self) -> bool:
        return not any(u.seviye == "error" for u in self.uyarilar)


def _tekil(degerler, ayrac="\n") -> str:
    gorulen = []
    for d in degerler:
        if d.strip() and d not in gorulen:
            gorulen.append(d)
    return ayrac.join(gorulen)


def _ria_yok(r: str) -> bool:
    return r.strip() in ("", "0")


def birlestir(satirlar: list[Satir]) -> tuple[list[Numune], list[Uyari]]:
    gruplar: dict[tuple, list[Satir]] = {}
    uyarilar, onceki_anahtar = [], {}
    for s in satirlar:
        anahtar = (s.grup, s.yas, s.sample_id)
        kume = (s.grup, s.yas)
        if anahtar in gruplar and onceki_anahtar.get(kume) != anahtar:
            satir_no = ", ".join(str(x.kaynak_satir) for x in gruplar[anahtar] + [s])
            uyarilar.append(Uyari("warning", f"{s.sample_id}: rows not next to each other (rows {satir_no}), merged"))
        gruplar.setdefault(anahtar, []).append(s)
        onceki_anahtar[kume] = anahtar

    numuneler = []
    for (grup, yas, sid), ss in gruplar.items():
        n = Numune(grup, yas, sid, ss,
                   ria=_tekil(x.ria for x in ss if not _ria_yok(x.ria)),
                   section=_tekil((x.section for x in ss), " & "),
                   chainage=_tekil(x.chainage for x in ss),
                   structure=_tekil(x.structure for x in ss),
                   element=_tekil(x.element for x in ss),
                   mix_no=_tekil(x.mix_no for x in ss),
                   sinif=_tekil(x.sinif for x in ss))
        n.shotcrete = any(k in x.element.lower() for x in ss for k in ("shotcrete", "slope"))
        n.no_ria = GRUP[grup].ria_var and not n.ria
        numuneler.append(n)
    return numuneler, uyarilar


def grout_yasi(n: Numune, test_tarihi: date) -> tuple[str, Uyari | None]:
    dokum = next((s.dokum for s in n.satirlar if s.dokum), None)
    fark = (test_tarihi - dokum).days if dokum else None
    if fark == 1:
        return "1d", None
    if fark == 7:
        return "7d", None
    return "?", Uyari("warning", f"{n.sample_id}: cannot tell 1 or 7 days, check")


def _sayi(sid: str):
    m = _SAYI.match(sid.strip())
    return (m[1], int(m[2]), len(m[2])) if m else None


def _id(on: str, sayi: int, genislik: int) -> str:
    return f"{on}{sayi:0{genislik}d}"


def seri_kontrol(ad: str, numuneler: list[Numune], onceki: tuple[date, list[str]] | None) -> list[Uyari]:
    parcalar = sorted({p for p in (_sayi(n.sample_id) for n in numuneler) if p}, key=lambda p: p[1])
    uy = []
    for a, b in zip(parcalar, parcalar[1:]):
        eksik = b[1] - a[1] - 1
        if eksik > EN_COK_EKSIK:
            uy.append(Uyari("warning", f"{ad}: {eksik} numbers missing between "
                                       f"{_id(*a)} and {_id(*b)}"))
        else:
            uy += [Uyari("warning", f"{ad}: {_id(a[0], k, a[2])} missing") for k in range(a[1] + 1, b[1])]
    if onceki and parcalar:
        dunku = [p for p in (_sayi(i) for i in onceki[1]) if p]
        if dunku:
            son = max(dunku, key=lambda p: p[1])
            ilk = parcalar[0]
            if son[1] + 1 != ilk[1]:
                uy.append(Uyari("warning", f"{ad}: previous day ({onceki[0]:%d.%m.%Y}) ended at {_id(*son)}, "
                                           f"today starts at {_id(*ilk)}"))
    return uy


def yas_adi(grup_kod: str, yas: str) -> str:
    return f"{GRUP[grup_kod].ad} {'1d/7d' if yas == '1or7' else yas}"


def hesapla(tarih: date, secim: dict[str, Path], gruplar: tuple[str, ...] | None = None) -> Sonuc:
    sonuc = Sonuc()
    for g in GRUPLAR:
        if gruplar is not None and g.kod not in gruplar:
            continue
        yol = secim.get(g.kod)
        if not yol:
            sonuc.uyarilar.append(Uyari("error", f"{g.ad}: no source file chosen"))
            continue
        o = kaynak.oku(Path(yol), g, tarih)
        sonuc.uyarilar += o.uyarilar
        numuneler, uy = birlestir(o.satirlar)
        sonuc.uyarilar += uy
        for yas in g.yaslar:
            ns = [n for n in numuneler if n.yas == yas]
            satir_sayisi = sum(n.set_sayisi for n in ns)
            ad = yas_adi(g.kod, yas)
            sonuc.uyarilar += seri_kontrol(ad, ns, o.onceki.get(yas))
            bos_id = sum(1 for u in o.uyarilar if f" {yas} row " in u.metin and "Sample ID is empty" in u.metin)
            if o.sayim.get(yas, 0) != satir_sayisi + bos_id:
                sonuc.uyarilar.append(Uyari("error", f"{ad}: count mismatch, {o.sayim.get(yas, 0)} source rows "
                                                     f"but {satir_sayisi} read"))
            sonuc.hesap.append(Hesap(ad, o.sayim.get(yas, 0), len(ns), satir_sayisi - len(ns),
                                     sum(n.no_ria for n in ns), not g.davete))
        for n in numuneler:
            if n.yas == "1or7":
                n.yas, u = grout_yasi(n, tarih)
                if u:
                    sonuc.uyarilar.append(u)
        sonuc.numuneler += numuneler
    return sonuc
