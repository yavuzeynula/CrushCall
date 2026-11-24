"""Not metni (Yavuz'un Not Defteri'ne yapıştırdığı biçim)."""
from datetime import date

from veri import Numune

_YAS_SIRA = {"1d": 0, "?": 1, "3d": 2, "7d": 3, "28d": 4}
_MADDE_GRUP = ("grout_e", "grout_s", "cmd", "tcon")   # madde işaretli satırların sırası (grout'lar birlikte)


def yas_yazisi(yas: str) -> str:
    return {"1d": "1day", "?": "1or7days"}.get(yas, yas.replace("d", "days"))


def _ekler(n: Numune) -> str:
    return (" SHOTCRETE" if n.shotcrete else "") + (" NO RIA" if n.no_ria else "") + \
           (f" {n.set_sayisi} SET" if n.set_sayisi > 1 else "")


def yap(numuneler: list[Numune], tarih: date) -> str:
    satirlar = []
    for yas in ("7d", "28d"):
        blok = [n for n in numuneler if n.grup == "concrete" and n.yas == yas]
        if blok:
            satirlar.append(f"///{yas_yazisi(yas)}-{tarih:%d.%m.%Y}\\\\\\")
            satirlar += [f"{i}. {n.sample_id}{_ekler(n)}" for i, n in enumerate(blok, 1)]
    madde = [n for n in numuneler if n.grup in _MADDE_GRUP]
    grup_sira = lambda n: 0 if n.grup.startswith("grout") else _MADDE_GRUP.index(n.grup)
    madde.sort(key=lambda n: (grup_sira(n), _YAS_SIRA.get(n.yas, 9)))   # sıralama kararlı: kaynak sırası korunur
    satirlar += [f"• {n.sample_id} {yas_yazisi(n.yas)}{_ekler(n)}" for n in madde]
    return "".join(s + "\n" for s in satirlar)
