"""Kullanıcı ayarları: ~/.crushcall/ayarlar.json (Windows'ta da ev klasörü)."""
import json
from pathlib import Path


def dosya() -> Path:
    return Path.home() / ".crushcall" / "ayarlar.json"


def oku() -> dict:
    d = dosya()
    if not d.exists():
        return {}
    try:
        return json.loads(d.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        d.replace(d.with_name(d.name + ".bozuk"))   # bozuk dosyayı sakla, temiz başla
        return {}


def yaz(ayarlar: dict) -> None:
    d = dosya()
    d.parent.mkdir(parents=True, exist_ok=True)
    gecici = d.with_name(d.name + ".tmp")
    gecici.write_text(json.dumps(ayarlar, ensure_ascii=False, indent=2), encoding="utf-8")
    gecici.replace(d)
