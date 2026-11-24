import cikti_not
from tests.fixtures import D25
from veri import Numune


def N(sid, grup="concrete", yas="7d", sh=False, set_=1, no_ria=False):
    return Numune(grup, yas, sid, [None] * set_, shotcrete=sh, no_ria=no_ria)


def test_25_11_ornegi():
    ns = ([N("G26-CON-1713"), N("G26-CON-1714", sh=True, set_=2)] +
          [N(f"G26-CON-{i}") for i in range(1715, 1720)] +
          [N(f"G26-CON-{i}", yas="28d") for i in range(1594, 1598)] +
          [N("G26-CON-1598", yas="28d", sh=True, set_=2), N("G26-CON-1599", yas="28d"), N("G26-CON-1600", yas="28d"),
           N("G26-GRT-0240", grup="grout_s", yas="28d"), N("G26-GRT-0254", grup="grout_s", yas="7d")])
    assert cikti_not.yap(ns, D25) == (
        "///7days-25.11.2026\\\\\\\n"
        "1. G26-CON-1713\n2. G26-CON-1714 SHOTCRETE 2 SET\n3. G26-CON-1715\n4. G26-CON-1716\n"
        "5. G26-CON-1717\n6. G26-CON-1718\n7. G26-CON-1719\n"
        "///28days-25.11.2026\\\\\\\n"
        "1. G26-CON-1594\n2. G26-CON-1595\n3. G26-CON-1596\n4. G26-CON-1597\n"
        "5. G26-CON-1598 SHOTCRETE 2 SET\n6. G26-CON-1599\n7. G26-CON-1600\n"
        "• G26-GRT-0254 7days\n• G26-GRT-0240 28days\n")


def test_bos_blok_yok():
    assert cikti_not.yap([N("G26-CON-1594", yas="28d")], D25) == "///28days-25.11.2026\\\\\\\n1. G26-CON-1594\n"
    assert cikti_not.yap([], D25) == ""


def test_cmd_tcon_ve_yas_yazimi():
    ns = [N("T-1", grup="tcon", yas="3d"), N("CMD-0123", grup="cmd", yas="28d"), N("G-1", grup="grout_e", yas="1d"),
          N("G-2", grup="grout_s", yas="?")]
    assert cikti_not.yap(ns, D25) == "• G-1 1day\n• G-2 1or7days\n• CMD-0123 28days\n• T-1 3days\n"


def test_no_ria_eki():
    assert cikti_not.yap([N("A-1", no_ria=True, sh=True, set_=3)], D25).splitlines()[1] == "1. A-1 SHOTCRETE NO RIA 3 SET"
