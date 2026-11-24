import os
import sys
import tkinter as tk

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for p in (os.path.join(ROOT, "src"), os.path.join(ROOT, "build"), ROOT):
    if p not in sys.path:
        sys.path.insert(0, p)

_KOK = []


@pytest.fixture
def tk_root():
    """Süreç başına TEK Tk (Windows'ta art arda Tk() ara sıra init.tcl okuyamıyor)."""
    if not _KOK:
        r = tk.Tk()
        r.withdraw()
        _KOK.append(r)
    r = _KOK[0]
    tk._default_root = r
    yield r
    for w in r.winfo_children():
        w.destroy()


@pytest.fixture(autouse=True)
def _ev_klasoru(tmp_path, monkeypatch):
    """Testler gerçek ~/.crushcall ayarına dokunmasın."""
    from pathlib import Path
    ev = tmp_path / "ev"
    ev.mkdir()
    monkeypatch.setattr(Path, "home", lambda: ev)
