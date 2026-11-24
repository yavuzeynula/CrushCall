"""assets/app.ico'yu amblemden uretir (ikon.app_icon)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
import ikon  # noqa: E402

(ROOT / "assets").mkdir(exist_ok=True)
ikon.app_icon(256).save(ROOT / "assets" / "app.ico",
                        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (256, 256)])
print("assets/app.ico")
