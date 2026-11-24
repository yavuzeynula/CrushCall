# pip-audit'i Windows sertifika deposuyla çalıştırır (ağ sertifikayı araya girip imzalıyor).
# Ardından kilitteki sürümler PyPI'de geri çekilmiş (yanked) mi diye bakar: pip-audit bunu
# görmüyor (01.11.2026: pikepdf 10.7.0 "bazı Python'larda çöküyor" diye çekilmişti).
# Açık ya da geri çekilmiş sürüm varsa çıkış kodu 1.
import json, re, sys, urllib.request, truststore
truststore.inject_into_ssl()
from pip_audit._cli import audit


def geri_cekilmis(dosyalar):
    bulunan = []
    for f in dosyalar:
        for ad, v in re.findall(r"^([A-Za-z0-9_.\-]+)==([^\s;\\]+)", open(f, encoding="utf-8").read(), re.M):
            d = json.load(urllib.request.urlopen(f"https://pypi.org/pypi/{ad}/{v}/json", timeout=30))
            if any(u.get("yanked") for u in d["urls"]):
                bulunan.append(f"{ad}=={v} ({d['urls'][0].get('yanked_reason') or 'neden yok'})")
    return bulunan


dosyalar = [sys.argv[i + 1] for i, a in enumerate(sys.argv[:-1]) if a in ("-r", "--requirement")]
sys.argv[0] = "pip-audit"
try:
    audit()
    kod = 0
except SystemExit as e:
    kod = e.code or 0
cekilen = geri_cekilmis(dosyalar)
for c in cekilen:
    print("PyPI'de geri çekilmiş sürüm:", c)
sys.exit(1 if (kod or cekilen) else 0)
