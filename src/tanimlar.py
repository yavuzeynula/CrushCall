"""Gruplar ve kaynak dosyalardaki sütun başlıkları.

Başlıklar eski Power Query sorgularından alındı (ornekler/eski-sorgular.m).
Karşılaştırma boşluk farkını saymaz (bkz. kaynak.normal).
"""
from dataclasses import dataclass


@dataclass(frozen=True)
class Grup:
    kod: str
    ad: str                          # ekranda görünen ad
    sayfa: str                       # Excel sayfa adı
    yaslar: dict[str, tuple[str, ...]]  # yaş -> kabul edilen tarih sütunu başlıkları (ilk bulunan)
    zorunlu: tuple[str, ...]         # bulunmazsa grup okunmaz
    nota: bool = True
    davete: bool = True
    ria_var: bool = True             # NO RIA kontrolü yapılır mı
    yer_var: bool = True             # Location + sağındaki 3 sütun okunur mu


def _beton(n: int) -> tuple[str, ...]:
    return (f"{n} days Compressive Strength EN12390-3",)


GRUPLAR: tuple[Grup, ...] = (
    Grup("concrete", "Concrete", "Concrete Comp Log",
         {"7d": _beton(7), "28d": _beton(28)},
         ("Sample ID", "RIA No:", "Location")),
    Grup("cmd", "CMD", "Concrete Comp Log",
         {"3d": _beton(3), "7d": _beton(7), "28d": _beton(28)},
         ("Sample ID", "Mix No", "Concrete Class"),
         ria_var=False, yer_var=False),
    Grup("grout_e", "Grout (Elastomeric)", "Grout Comp Log",
         {"1d": ("1 day Compressive Strength EN 196-1",),
          "28d": ("28 days Compressive Strength EN 196-1", "28 day Compressive Strength EN 196-1")},
         ("Sample ID", "RIA No:", "Location")),
    Grup("grout_s", "Grout (Slope)", "Grout Comp Log",
         {"1or7": ("1-day or 7-days Compressive Strength EN 196-1",),
          "28d": ("28-days Compressive Strength EN 196-1",)},
         ("Sample ID", "RIA No:", "Location", "Pouring Date")),
    Grup("tcon", "T-CON", "Concrete Comp Log",
         {"3d": _beton(3), "7d": _beton(7), "28d": _beton(28)},
         ("Sample ID",),
         davete=False, ria_var=False, yer_var=False),
)

GRUP = {g.kod: g for g in GRUPLAR}

# Davet listesi sabitleri (örnek dosyadaki gibi)
LAB = "BEJV LAB"
YONTEM = "MKC EN 12390-3:2019"
SAAT = "09:00"
CMD_SAAT = "11:00"
