"""Ortak veri tipleri."""
from dataclasses import dataclass, field
from datetime import date


@dataclass
class Satir:
    """Kaynak Excel'den okunan tek satır."""
    grup: str            # "concrete" | "cmd" | "grout_e" | "grout_s" | "tcon"
    yas: str             # "1d" "3d" "7d" "28d" ya da "1or7" (grout_s; islem çözer)
    sample_id: str
    test_tarihi: date
    kaynak_satir: int    # Excel satır numarası
    ria: str = ""
    mix_no: str = ""
    sinif: str = ""
    dokum: date | None = None
    section: str = ""
    chainage: str = ""
    structure: str = ""
    element: str = ""


@dataclass
class Numune:
    """Aynı Sample ID'li satırların birleşmiş hali."""
    grup: str
    yas: str
    sample_id: str
    satirlar: list[Satir]
    ria: str = ""
    section: str = ""
    chainage: str = ""
    structure: str = ""
    element: str = ""
    mix_no: str = ""
    sinif: str = ""
    shotcrete: bool = False
    no_ria: bool = False

    @property
    def set_sayisi(self) -> int:
        return len(self.satirlar)


@dataclass
class Uyari:
    seviye: str          # "error" | "warning"
    metin: str           # İngilizce; ekranda olduğu gibi görünür


@dataclass
class Okuma:
    """Bir kaynak dosyanın bir tarih için okunması."""
    satirlar: list[Satir] = field(default_factory=list)
    uyarilar: list[Uyari] = field(default_factory=list)
    sayim: dict[str, int] = field(default_factory=dict)          # yaş -> tarihi eşleşen hücre sayısı
    onceki: dict[str, tuple[date, list[str]]] = field(default_factory=dict)  # yaş -> (önceki test günü, Sample ID'ler)
