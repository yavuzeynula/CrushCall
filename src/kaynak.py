"""Kaynak kırım log'larını okur. Dosyalara asla yazmaz.

Düzen (eski Power Query ile aynı): ilk 3 satır atlanır, 4. satır başlık,
5. satırdan veri. Tarih sütunundaki değer = testin yapılacağı gün.
Okunamayan hiçbir şey sessiz geçmez: her biri Uyari olarak döner.
"""
import re
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

from tanimlar import Grup
from veri import Okuma, Satir, Uyari

BASLIK_SATIRI = 4
EN_COK_ORNEK = 5          # aynı türden uyarıların kaçı tek tek yazılır
_EXCEL_SIFIR = date(1899, 12, 30)
_METIN_TARIH = re.compile(r"^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$")


def normal(x) -> str:
    """Başlık karşılaştırması için: boşluk ve satır sonu farkını yok sayar."""
    return " ".join(str(x).split()) if x is not None else ""


def metin(v) -> str:
    """Hücre değerini olduğu gibi metne çevirir (iç boşluk ve satır sonları kalır)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def tarih_coz(v) -> tuple[date | None, str]:
    """(tarih, durum): durum "ok" | "bos" | "hata"."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None, "bos"
    if isinstance(v, datetime):
        return v.date(), "ok"
    if isinstance(v, date):
        return v, "ok"
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 <= v <= 80000:
        return _EXCEL_SIFIR + timedelta(days=int(v)), "ok"   # Excel seri sayısı (Power Query de sayıyordu)
    if isinstance(v, str):
        m = _METIN_TARIH.match(v.strip())
        if m:
            try:
                return date(int(m[3]), int(m[2]), int(m[1])), "ok"   # yalnız gün/ay/yıl
            except ValueError:
                pass
    return None, "hata"


def _al(satir, i):
    return satir[i] if i is not None and i < len(satir) else None


def oku(yol: Path, grup: Grup, tarih: date) -> Okuma:
    o = Okuma()
    yol = Path(yol)
    hata = lambda m: o.uyarilar.append(Uyari("error", f"{grup.ad}: {m}"))
    if not yol.is_file():
        hata(f"file not found: {yol}")
        return o
    try:
        wb = load_workbook(yol, read_only=True, data_only=True)
    except Exception as e:  # bozuk / desteklenmeyen dosya
        hata(f"cannot open {yol.name}: {e}")
        return o
    try:
        bekleyen_formul = _tara(wb, grup, tarih, o, yol)
    finally:
        wb.close()
    if bekleyen_formul:
        _formul_kontrol(yol, grup, bekleyen_formul, o)
    return o


def _tara(wb, grup: Grup, tarih: date, o: Okuma, yol: Path) -> dict[int, list[tuple[int, str]]]:
    hata = lambda m: o.uyarilar.append(Uyari("error", f"{grup.ad}: {m}"))
    if grup.sayfa not in wb.sheetnames:
        hata(f"sheet '{grup.sayfa}' not found in {yol.name}")
        return {}
    satirlar = wb[grup.sayfa].iter_rows(values_only=True)
    baslik = next((s for n, s in enumerate(satirlar, 1) if n == BASLIK_SATIRI), None) or ()
    yer = {}
    for i, h in enumerate(baslik):
        yer.setdefault(normal(h), i) if normal(h) else None

    for ad in grup.zorunlu:
        if ad not in yer:
            hata(f"header '{ad}' not found in row {BASLIK_SATIRI} of {yol.name}")
    tarih_sut = {}
    for yas, adlar in grup.yaslar.items():
        i = next((yer[normal(a)] for a in adlar if normal(a) in yer), None)
        if i is None:
            hata(f"header '{adlar[0]}' not found in row {BASLIK_SATIRI} of {yol.name}")
        tarih_sut[yas] = i
    if any(u.seviye == "error" for u in o.uyarilar):
        return {}

    sid_i, loc = yer["Sample ID"], yer.get("Location")
    o.sayim = {yas: 0 for yas in tarih_sut}
    bozuk, bekleyen = [], {}
    for n, satir in enumerate(satirlar, BASLIK_SATIRI + 1):
        sid = metin(_al(satir, sid_i)).strip()
        for yas, i in tarih_sut.items():
            v = _al(satir, i)
            d, durum = tarih_coz(v)
            if durum == "bos":
                if v is None and sid:
                    bekleyen.setdefault(n, []).append((i, yas))
                continue
            if durum == "hata":
                bozuk.append(f"{grup.ad} {yas} row {n}: date not readable: '{v}'")
                continue
            if d == tarih:
                o.sayim[yas] += 1
                if not sid:
                    hata(f"{yas} row {n}: date matches but Sample ID is empty")
                    continue
                o.satirlar.append(_satir(grup, yas, sid, tarih, n, satir, yer, loc))
            elif d < tarih:
                onceki = o.onceki.get(yas)
                if sid and (onceki is None or d > onceki[0]):
                    o.onceki[yas] = (d, [sid])
                elif sid and d == onceki[0]:
                    onceki[1].append(sid)

    for m in bozuk[:EN_COK_ORNEK]:
        o.uyarilar.append(Uyari("warning", m))
    if len(bozuk) > EN_COK_ORNEK:
        o.uyarilar.append(Uyari("warning", f"{grup.ad}: {len(bozuk) - EN_COK_ORNEK} more unreadable date cells"))
    return bekleyen


def _satir(grup, yas, sid, tarih, n, satir, yer, loc) -> Satir:
    s = Satir(grup=grup.kod, yas=yas, sample_id=sid, test_tarihi=tarih, kaynak_satir=n,
              ria=metin(_al(satir, yer.get("RIA No:"))),
              mix_no=metin(_al(satir, yer.get("Mix No"))),
              sinif=metin(_al(satir, yer.get("Concrete Class"))),
              dokum=tarih_coz(_al(satir, yer.get("Pouring Date")))[0])
    if grup.yer_var and loc is not None:
        s.section, s.chainage, s.structure, s.element = (metin(_al(satir, loc + k)) for k in range(4))
    return s


def _formul_kontrol(yol: Path, grup: Grup, bekleyen: dict, o: Okuma) -> None:
    """Değeri boş görünen tarih hücreleri formül mü? Formülün kayıtlı değeri yoksa
    o gün bilinemez: veri kaybı riski -> hata."""
    wb = load_workbook(yol, read_only=True, data_only=False)
    bulunan = []
    try:
        for n, satir in enumerate(wb[grup.sayfa].iter_rows(values_only=True), 1):
            for i, yas in bekleyen.get(n, ()):
                v = _al(satir, i)
                if isinstance(v, str) and v.startswith("="):
                    bulunan.append(f"{grup.ad} {yas} row {n}: formula has no saved value "
                                   "(open and save the file in Excel)")
    finally:
        wb.close()
    for m in bulunan[:EN_COK_ORNEK]:
        o.uyarilar.append(Uyari("error", m))
    if len(bulunan) > EN_COK_ORNEK:
        o.uyarilar.append(Uyari("error", f"{grup.ad}: {len(bulunan) - EN_COK_ORNEK} more formulas without saved value"))
