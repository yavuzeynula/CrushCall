"""Davet listesi: satırlar, .xlsx (örnek dosyayla aynı biçim), panoya biçimli kopya."""
import html as _html
from copy import copy
from datetime import date, time
import tempfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from tanimlar import CMD_SAAT, LAB, SAAT, YONTEM
from veri import Numune

SABLON = Path(__file__).resolve().parent / "sablon" / "davet_sablonu.xlsx"
ILK_SATIR = 13
SARI, YESIL = "FFFF00", "92D050"
_YAS = {"1d": 0, "?": 1, "3d": 2, "7d": 3, "28d": 4}
_GRUP = {"concrete": 0, "cmd": 1, "grout_e": 2, "grout_s": 2}


class KayitHatasi(Exception):
    pass


def _test_adi(n: Numune) -> str:
    return ("Grout \n" if n.grup.startswith("grout") else "") + f"Compressive Strength ({n.yas}) "


def satirlar(numuneler: list[Numune], tarih: date) -> list[list]:
    secilen = [n for n in numuneler if n.grup in _GRUP and not n.no_ria]
    secilen.sort(key=lambda n: (_GRUP[n.grup], _YAS.get(n.yas, 9)))   # kararlı: kaynak sırası korunur
    gun = f"{tarih:%d.%m.%Y}"
    out = []
    for i, n in enumerate(secilen, 1):
        if n.grup == "cmd":
            out.append([i, LAB, _test_adi(n), gun, "-", "-", "1", "11+500\nQC Lab", "-",
                        f"Concrete Mix design  #{n.mix_no} {n.sinif}", CMD_SAAT])
        else:
            out.append([i, LAB, _test_adi(n), gun, n.ria, YONTEM, n.section, n.chainage, n.structure, n.element, SAAT])
    return out


def _yesil_mi(satir: list) -> bool:
    return satir[2].startswith("Grout") or satir[5] == "-"


def dosya_adi(tarih: date) -> str:
    return f"Laboratory_Daily_Notification_Invitation_List_{tarih:%d.%m.%Y}.xlsx"


def kopya_dosyasi(satirlar_: list[list], tarih: date) -> Path:
    """Copy rows için geçici dosya. Adı yalnız tarih: kullanıcının açık davet dosyasıyla
    aynı adı taşımasın (Excel aynı adlı iki kitabı açamaz)."""
    return kaydet(satirlar_, tarih, Path(tempfile.mkdtemp(prefix="CrushCall_")), f"{tarih:%d.%m.%Y}.xlsx")


def kaydet(satirlar_: list[list], tarih: date, klasor: Path, ad: str | None = None) -> Path:
    wb = load_workbook(SABLON)
    ws = wb.active
    ws.title = f"{tarih:%d.%m.%Y}"
    ornek = {c: ws.cell(ILK_SATIR, c) for c in range(3, 14)}
    stil = {c: (copy(h.font), copy(h.border), copy(h.fill), copy(h.alignment), h.number_format) for c, h in ornek.items()}
    for i, s in enumerate(satirlar_):
        r = ILK_SATIR + i
        for k, deger in enumerate(s):
            c = 3 + k
            hucre = ws.cell(r, c)
            font, kenar, dolgu, hiza, bicim = stil[c]
            hucre.font, hucre.border, hucre.fill, hucre.alignment = copy(font), copy(kenar), copy(dolgu), copy(hiza)
            hucre.number_format = bicim
            if c == 13:
                deger = time(*map(int, deger.split(":")))
            hucre.value = deger
        if _yesil_mi(s):
            ws.cell(r, 5).fill = PatternFill("solid", fgColor="FF" + YESIL)
        satir_sayisi = max(str(v).count("\n") + 1 for v in s)
        ws.row_dimensions[r].height = max(30, 15 * satir_sayisi)
    yol = Path(klasor) / (ad or dosya_adi(tarih))
    try:
        wb.save(yol)
    except PermissionError as e:
        raise KayitHatasi(f"{yol.name} is open in another program; close it in Excel and try again") from e
    return yol


def html(satirlar_: list[list]) -> str:
    """Excel'e yapıştırınca biçimiyle gelen tablo. Başlık yok, yalnız veri satırları."""
    tel = "border:.5pt solid windowtext;font-family:Calibri,sans-serif;font-size:11pt;color:black;" \
          "text-align:center;vertical-align:middle;white-space:normal;"
    br = '<br style="mso-data-placement:same-cell">'
    satir_html = []
    for s in satirlar_:
        hucreler = []
        for k, v in enumerate(s):
            renk = "" if k == 0 else f"background:#{YESIL if k == 2 and _yesil_mi(s) else SARI};"
            hucreler.append(f'<td style="{tel}{renk}">{_html.escape(str(v)).replace(chr(10), br)}</td>')
        satir_html.append("<tr>" + "".join(hucreler) + "</tr>")
    return "<table style=\"border-collapse:collapse\">" + "".join(satir_html) + "</table>"


def tsv(satirlar_: list[list]) -> str:
    """Düz metin yedeği: Excel'in anladığı sekmeli metin (çok satırlı hücre tırnakta)."""
    def hucre(v):
        v = str(v)
        return '"' + v.replace('"', '""') + '"' if any(c in v for c in '\n\t"') else v
    return "\r\n".join("\t".join(hucre(v) for v in s) for s in satirlar_)
