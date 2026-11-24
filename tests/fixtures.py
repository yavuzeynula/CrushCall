"""Testler için küçük kaynak Excel üretici (gerçek log düzenini taklit eder)."""
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

BETON_BASLIK = ["RIA No:", "Sample ID", "Mix No", "Supplier", "Pouring Date", "Sampling Location",
                "Location", None, None, None,
                "7 days Compressive Strength EN12390-3", "28 days Compressive Strength EN12390-3"]


def kaynak_yap(yol: Path, basliklar: list, satirlar: list[list], sayfa: str = "Concrete Comp Log",
               ust_bos: int = 3) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = sayfa
    for _ in range(ust_bos):
        ws.append([None])
    ws.append(basliklar)
    for s in satirlar:
        ws.append(s)
    wb.save(yol)
    return yol


def beton_satir(ria, sid, t7=None, t28=None, dokum=None, sec="2", ch="1+000", st="Bridge", el="Pile"):
    return [ria, sid, "CMD-7301", "Supplier A", dokum, "Site", sec, ch, st, el, t7, t28]


def gun(g: int, a: int = 11) -> datetime:
    return datetime(2026, a, g)


D25 = date(2026, 11, 25)
