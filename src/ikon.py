"""Uygulama amblemi: koyu kare üstünde çatlamış beton küp (kırım)."""
from PIL import Image, ImageDraw

KAT = 4   # büyük çiz, küçült (kenar yumuşak olsun); yalnız açılışta bir kez


def app_icon(n: int = 256) -> Image.Image:
    b = n * KAT
    im = Image.new("RGBA", (b, b), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, b - 1, b - 1), b * 0.2, fill=(8, 21, 25, 255))
    cx, ust, alt, yan = b * 0.5, b * 0.2, b * 0.82, b * 0.28
    orta = b * 0.42
    tepe = [(cx, ust), (cx + yan, ust + b * 0.11), (cx, orta + b * 0.0), (cx - yan, ust + b * 0.11)]
    sol = [(cx - yan, ust + b * 0.11), (cx, orta), (cx, alt), (cx - yan, alt - b * 0.11)]
    sag = [(cx + yan, ust + b * 0.11), (cx, orta), (cx, alt), (cx + yan, alt - b * 0.11)]
    d.polygon(tepe, fill=(244, 186, 140, 255))
    d.polygon(sol, fill=(232, 147, 90, 255))
    d.polygon(sag, fill=(184, 98, 42, 255))
    w = round(b * 0.03)
    d.line([(cx + yan * 0.55, orta - b * 0.02), (cx + yan * 0.35, b * 0.55), (cx + yan * 0.7, b * 0.63),
            (cx + yan * 0.45, b * 0.75)], fill=(82, 224, 187, 255), width=w, joint="curve")
    return im.resize((n, n), Image.LANCZOS)


def set_icon(kok) -> None:
    try:
        from PIL import ImageTk
        img = ImageTk.PhotoImage(app_icon(64), master=kok)
        kok._crushcall_ikon = img          # çöpe gitmesin
        kok.iconphoto(True, img)
    except Exception:
        pass
