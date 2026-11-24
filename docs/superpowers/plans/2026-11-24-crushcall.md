# CrushCall Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (Yavuz'un genel kuralı: "paralel çalış" denmedikçe tek oturumda). Steps use checkbox (`- [ ]`) syntax.

**Goal:** Kaynak kırım Excel'lerinden seçilen tarihte testi olan numuneleri çekip not metni ve davet listesi Excel'i üreten, lisanslı ve kendini güncelleyen Windows masaüstü programı.

**Architecture:** Saf Python çekirdek (`src/` içinde okuma → işleme → çıktı, Tk'siz, testli) + ince Tk ekranı. Exe kabuğu, lisans, güncelleme ve yayın CheckList'ten kopyalanıp adları değiştirilir.

**Tech Stack:** Python 3.12, tkinter, openpyxl 3.1.5, pillow, cryptography, pyinstaller, pytest + pytest-xdist.

**Spec:** `docs/superpowers/specs/2026-11-24-crushcall-design.md`

## Global Constraints

- Ekrandaki bütün yazılar İngilizce; kod yorumları Türkçe olabilir.
- Yol: `pathlib`; dosya okuma/yazma `encoding="utf-8"`; ev klasörü `Path.home()`; Mac'e özgü her şeyin Windows karşılığı.
- Kaynak dosyalar yalnız okunur (`load_workbook(read_only=True, data_only=True)`); hiçbir hata sessiz geçmez.
- Davet listesi örnek dosyayla birebir: `ornekler/Laboratory_Daily_Notification_Invitation_List_25.11.2026.xlsx` (11 sütun C–M, veri 13. satırdan, sarı `FFFFFF00`, CMD/Grout Test hücresi `FF92D050`, yöntem her satırda `MKC EN 12390-3:2019`, CMD'de `-`; saat 09:00, CMD 11:00).
- Not biçimi: `ornekler/not-ornegi.jpg`; `2 SET`, `SHOTCRETE`, `NO RIA` ekleri.
- NO RIA Beton/Grout satırı davete girmez; T-CON yalnız nota.
- Paketler kilitli: `requirements.in` → `pip-compile --generate-hashes --strip-extras --allow-unsafe` → `~/tools/tedarik/yas.py -r requirements.txt` → `tara.py` → `pip install --require-hashes`.
- Testler paralel: `pytest -n auto --dist worksteal`; sıralı koşu yok.
- `ornekler/` ve `prototip/` gerçek şirket verisi içerir: `.gitignore`'da, depoya girmez (GitHub deposu güncelleme için açık olacak).
- Commit yalnız Yavuz isteyince.

## Review Focus

1. Kaynakta başlık satırı 4. satırdan farklı yerdeyse (biri üste satır eklediyse) → "header not found" hatası, boş liste değil. Test: Task 2 `test_baslik_kaymis_hata`.
2. Tarih hücresi metin `28/10/2026` ya da `28.10.2026` → tarih sayılır; `10/28/2026` → uyarı. Test: Task 2 `test_metin_tarih`.
3. Aynı Sample ID ardışık değil de araya başka satır girmiş → yine birleşir ve uyarı verir. Test: Task 3 `test_ardisik_olmayan_double`.
4. Formüllü hücre, değeri hesaplanmamış (None) → uyarı. Test: Task 2 `test_formul_degersiz`.
5. Kaynak dosya başka biri tarafından açıkken/kilitliyken → okunur ya da açık hata mesajı. Test: Task 2 `test_dosya_yok_hata` (kilit Windows'ta elle denenir).

---

## Dosya düzeni

```
launcher.py, license_client.py, license_secret.txt(gitignore)   ← CheckList'ten (Task 8)
src/
  main.py        Tk açılışı + selftest()                        (Task 7)
  version.txt
  tanimlar.py    gruplar, sütun başlıkları, sabit metinler        (Task 2)
  kaynak.py      Excel okuma → Satir listesi + Uyari             (Task 2)
  islem.py       birleştirme, işaret, grout yaşı, seri, hesap     (Task 3)
  cikti_not.py   not metni                                        (Task 4)
  cikti_davet.py davet satırları, .xlsx, HTML/TSV pano            (Task 5)
  pano.py        Windows/Mac biçimli pano                         (Task 5)
  ayar.py        ~/.crushcall/ayarlar.json                        (Task 6)
  ekran.py       Tk C düzeni                                      (Task 7)
  updater.py     CheckList'ten                                    (Task 8)
  sablon/davet_sablonu.xlsx                                       (Task 5)
build/ publish.py, CrushCall.spec, installer.iss, make_icon.py    (Task 8)
tests/  her modülün testi + fixtures.py (örnek kaynak üretici)
araclar/karsilastir.py  geçmiş günlerle kabul testi               (Task 9)
```

## Ortak veri tipleri (`src/veri.py`, Task 2'de tanımlanır)

```python
@dataclass
class Satir:            # kaynaktan okunan tek satır
    grup: str           # "concrete" | "cmd" | "grout_e" | "grout_s" | "tcon"
    yas: str            # "1d" "3d" "7d" "28d" | "1or7" (grout_s, Task 3 çözer)
    sample_id: str
    ria: str; mix_no: str; sinif: str
    dokum: date | None; test_tarihi: date
    section: str; chainage: str; structure: str; element: str
    kaynak_satir: int   # Excel satır numarası (uyarılar için)

@dataclass
class Numune:           # birleştirilmiş
    grup: str; yas: str; sample_id: str
    satirlar: list[Satir]
    ria: str; section: str; chainage: str; structure: str; element: str
    mix_no: str; sinif: str
    shotcrete: bool; no_ria: bool
    @property set_sayisi -> int   # len(satirlar)

@dataclass
class Uyari:
    seviye: str          # "error" | "warning"
    metin: str           # İngilizce, ekranda olduğu gibi
```

---

### Task 1: Proje iskeleti ve paketler

**Files:** Create `requirements.in`, `requirements.txt`, `pytest.ini`, `.gitignore`, `src/version.txt` (`0.1.0`), `tests/conftest.py`

- [ ] `requirements.in`: `openpyxl==3.1.5`, `pillow==13.0.0`, `cryptography==50.0.2`, `pyinstaller==6.19.0`, `pytest==9.0.3`, `pytest-xdist` (LabTrack'in test kilidindeki sürüm).
- [ ] Kilitle → `yas.py` → `tara.py` → `.venv` içine `--require-hashes` kur. Beklenen: üçü de "temiz".
- [ ] `pytest.ini`: `testpaths = tests`, `addopts = -n auto --dist worksteal`, `pythonpath = src .`
- [ ] `.gitignore`: CheckList'inki + `ornekler/`, `prototip/`, `.venv/`.
- [ ] `.venv/bin/pytest` → "no tests ran" (iskelet çalışıyor).

### Task 2: Tanımlar ve kaynak okuma

**Files:** Create `src/veri.py`, `src/tanimlar.py`, `src/kaynak.py`, `tests/fixtures.py`, `tests/test_kaynak.py`

**Interfaces:**
- Produces: `tanimlar.GRUPLAR: list[Grup]`; `Grup(kod, ad, sayfa, yaslar: dict[str, list[str]], nota: bool, davete: bool, ria_var: bool)` — `yaslar` yaş → kabul edilen başlık adları listesi (ilk bulunan kullanılır; grout_e 28d için `["28 days Compressive Strength EN 196-1", "28 day Compressive Strength EN 196-1"]`).
- Produces: `kaynak.oku(yol: Path, grup: Grup, tarih: date) -> tuple[list[Satir], list[Uyari], dict[str, int]]` — son değer yaş → bu tarihe eşit tarih hücresi sayısı (bağımsız sayım, Task 3 hesap kontrolü).
- Produces: `kaynak.onceki_gun(yol, grup, yas, tarih) -> tuple[date | None, list[str]]` — o sütunda tarihten önceki en yakın test günü ve o günün Sample ID'leri (kaynak sırasıyla).
- Produces: `fixtures.kaynak_yap(yol, sayfa, basliklar, satirlar, ust_bos=3)` — 3 boş satır + başlık + satırlar; `Location` başlığının sağındaki 3 sütun boş başlıklı.

Kurallar: sayfa adı `grup.sayfa`; başlık satırı = 4. satır; `Location` sütunu ve sağındaki 3 sütun = section, chainage, structure, element; hücre metinleri `str(...).strip()` değil — **iç boşluklar ve satır sonları korunur**, yalnız `None` → `""`; sayı hücresi `2` → `"2"` (float `2.0` ise `"2"`). Tarih: `datetime`/`date` doğrudan; metin `gg/aa/yyyy` veya `gg.aa.yyyy`; başka dolu değer → `warning` "Concrete row 812: date not readable: 'x'".

- [ ] Testler (`tests/test_kaynak.py`), her biri `fixtures.kaynak_yap` ile:
  - `test_tarihe_esit_satirlar_gelir`: 3 satırdan tarihi 25.11.2026 olan 2'si döner, alanlar doğru (`section` = Location sütunu, `element` = 3 sağı).
  - `test_metin_tarih`: `"25/11/2026"` ve `"25.11.2026"` alınır; `"11/25/2026"` → warning, satır alınmaz.
  - `test_baslik_kaymis_hata`: başlık 5. satırda → `error` "header '7 days Compressive Strength EN12390-3' not found", satır yok.
  - `test_grout_e_28_iki_yazim`: `"28 day ..."` başlığıyla da okur.
  - `test_formul_degersiz`: tarih hücresinde formül var, değer yok → warning.
  - `test_dosya_yok_hata`: olmayan yol → `error` "Concrete: file not found: ...".
  - `test_coklu_satir_korunur`: `"A\nB"` hücresi aynen gelir.
  - `test_onceki_gun`: 24.11'de 1711,1712; 22.11'de 1709 → `(24.11, ["…1711","…1712"])`.
- [ ] `.venv/bin/pytest tests/test_kaynak.py` → FAIL (modül yok).
- [ ] `veri.py`, `tanimlar.py`, `kaynak.py` yaz (yalnız openpyxl, `read_only=True, data_only=True`; formül kontrolü için ikinci açılış `data_only=False` yalnız değeri None olan tarih hücrelerinde).
- [ ] Testler PASS.

### Task 3: İşleme (birleştirme, işaret, grout yaşı, seri, hesap)

**Files:** Create `src/islem.py`, `tests/test_islem.py`

**Interfaces:**
- Consumes: `Satir`, `Uyari`, `kaynak.onceki_gun`.
- Produces: `islem.birlestir(satirlar: list[Satir]) -> tuple[list[Numune], list[Uyari]]` — grup+yaş+sample_id aynı olanlar tek Numune; alanlar: farklı değerler ilk görülme sırasıyla `"\n"` ile, `section` ` & ` ile; sıra = ilk görülme.
- Produces: `islem.grout_yasi(n: Numune, test_tarihi: date) -> tuple[str, Uyari | None]` — fark 1 → `"1d"`, 7 → `"7d"`, aksi/döküm yok → `"?"` + warning "G26-GRT-0254: cannot tell 1 or 7 days, check".
- Produces: `islem.seri_kontrol(ad: str, numuneler: list[Numune], onceki: list[str]) -> list[Uyari]` — Sample ID sonundaki sayı; eksikler "Concrete 7d: 1716 missing"; devam "Concrete 7d: previous day (24.11.2026) ended at 1712, today starts at 1715".
- Produces: `islem.hesapla(tarih, secim: dict[str, Path]) -> Sonuc` — `Sonuc(numuneler: list[Numune], uyarilar: list[Uyari], hesap: list[Hesap])`; `Hesap(ad, kaynak_satir, numune, birlesen, no_ria, sadece_not)`; `kaynak_satir != numune + birlesen` ya da `kaynak_satir != bağımsız sayım` → `error`, `Sonuc.tamam = False`.

- [ ] Testler:
  - `test_double_birlesir`: 1714 iki satır, RIA farklı → tek Numune, `ria == "R1\nR2"`, `set_sayisi == 2`.
  - `test_ardisik_olmayan_double`: 1714, 1715, 1714 → 1714 tek Numune + warning "G26-CON-1714: rows not next to each other".
  - `test_section_birlesir`: "2" ve "1" → `"2 & 1"`; aynıysa tek.
  - `test_shotcrete`: element "Slope 1" → `shotcrete`; "SHOTCRETE for" → `shotcrete`.
  - `test_no_ria`: ria `""` ve `"0"` → `no_ria`; cmd ve tcon grubunda hiç `no_ria` değil.
  - `test_grout_yasi`: döküm 24.11 test 25.11 → "1d"; 18.11 → "7d"; 20.11 → "?" + warning.
  - `test_seri_eksik`, `test_seri_devam_ok`, `test_seri_devam_kopuk`, `test_seri_sayisiz_id_atlanir`.
  - `test_hesap_tutmazsa_hata`.
- [ ] FAIL → yaz → PASS.

### Task 4: Not metni

**Files:** Create `src/cikti_not.py`, `tests/test_cikti_not.py`

**Interfaces:** `cikti_not.yap(numuneler: list[Numune], tarih: date) -> str`

Sıra: Beton 7d bloğu, Beton 28d bloğu (`///7days-25.11.2026\\\`, numaralı), sonra madde işaretliler: grout (1d,7d,28d), CMD, T-CON (`• G26-GRT-0254 7days`). Yaş yazımı: `1day`, `3days`, `7days`, `28days`. Ekler sırası: ` SHOTCRETE`, ` NO RIA`, ` 2 SET`. Boş blok yazılmaz. Satır sonu `\n`.

- [ ] `test_25_11_ornegi`: 25.11 verisi (fixture) → beklenen metin birebir (ekran görüntüsündeki + 1714/1598 için ` 2 SET`, grout 0254 `7days`).
- [ ] `test_bos_blok_yok`, `test_cmd_id_aynen` (`• CMD-0123 28days`), `test_no_ria_eki`.
- [ ] FAIL → yaz → PASS.

### Task 5: Davet listesi (.xlsx + biçimli pano)

**Files:** Create `src/cikti_davet.py`, `src/pano.py`, `src/sablon/davet_sablonu.xlsx`, `tests/test_cikti_davet.py`

**Interfaces:**
- `cikti_davet.satirlar(numuneler, tarih) -> list[list[str]]` — 11 hücre: `#`, `BEJV LAB`, test, `gg.aa.yyyy`, ria, yöntem, section, chainage, structure, element, saat. Test: `Compressive Strength (7d) ` (örnekteki gibi sonda boşluk), grout `Grout \nCompressive Strength (7d) `; CMD: ria/yöntem/structure `-`, section `1`, chainage `11+500\nQC Lab`, element `Concrete Mix design  #<mix_no> <sinif>`, saat `11:00`. NO RIA ve T-CON yok. Sıra: concrete 7d, 28d, cmd, grout.
- `cikti_davet.kaydet(satirlar, tarih, klasor: Path) -> Path` — şablondan; sayfa adı ve dosya adı `Laboratory_Daily_Notification_Invitation_List_<gg.aa.yyyy>.xlsx`; satır biçimi şablonun 13. satırından kopya; yeşil Test hücresi; satır yüksekliği = 15 × en çok satır sayısı.
- `cikti_davet.html(satirlar) -> str`, `cikti_davet.tsv(satirlar) -> str` — yalnız veri satırları (başlık yok). HTML'de satır sonu `<br style="mso-data-placement:same-cell">`, sarı/yeşil dolgu, ince siyah kenarlık, Calibri 11, ortalı.
- `pano.kopyala(html: str, metin: str, tk_kok) -> None` — Windows: `ctypes` ile `RegisterClipboardFormatW("HTML Format")` + `CF_UNICODETEXT`; Mac: `osascript` ile HTML + düz metin; başarısızsa Tk panosuna düz metin.

Şablon: örnek dosyanın kopyası, 13–28. satırların değerleri silinmiş. İlk adım ölçüm: openpyxl ile aç-kaydet sonrası resim sayısı (`ws._images`) 3 mü? Değilse resimleri `xl/media`'dan çıkarıp `drawing1.xml`'deki çapalara yeniden ekle.

- [ ] Testler:
  - `test_satirlar_25_11`: 16 satır; her hücre örnek dosyanın C13:M28 değerleriyle aynı (örnek dosyadan okunarak karşılaştırılır; saat hücresi `0.375` ↔ `09:00` eşlenir; F sütunu formül yerine tarih metni; `\r\n` ↔ `\n` eşlenir).
  - `test_no_ria_ve_tcon_girmez`, `test_cmd_satiri` (ekran görüntüsündeki 13. satır).
  - `test_kaydet_bicim`: kaydedilen dosyada C13 dolgu `FFFFFF00`, grout E hücresi `FF92D050`, kenarlık ince, sayfa adı tarih, resim sayısı 3.
  - `test_html_baslik_yok`, `test_tsv_cok_satirli_hucre_tirnakli`.
- [ ] FAIL → yaz → PASS. Windows'ta yapıştırma göz kontrolü Task 10'da.

### Task 6: Ayarlar

**Files:** Create `src/ayar.py`, `tests/test_ayar.py`

**Interfaces:** `ayar.oku() -> dict`, `ayar.yaz(d: dict) -> None` — `Path.home()/".crushcall"/"ayarlar.json"`, utf-8; anahtarlar `kaynaklar: {grup_kod: yol}`, `kayit_klasoru`, `tema` (`"dark"`/`"light"`). Bozuk dosya → boş ayar + yedek `ayarlar.json.bozuk`.

- [ ] `test_yaz_oku`, `test_bozuk_dosya` (ev klasörü `monkeypatch` ile tmp) → FAIL → yaz → PASS.

### Task 7: Ekran (Tk, C düzeni)

**Files:** Create `src/ekran.py`, `src/main.py`, `tests/test_ekran.py`

**Interfaces:** `main.main()` pencereyi açar; `main.selftest() -> int` pencere açmadan modülleri yükler, şablonu açar, 0 döner. `ekran.Pencere(kok: tk.Tk)`.

Düzen (spec "Ekran"): üst bant (CrushCall, `Test date ‹ gg.aa.yyyy ›`, `Pull`), sol 0,382 kontrol (grup · sayı · durum, altında `"8 source rows → 7 samples · 1 merged"`, uyarılar kırmızı/turuncu), sağ 0,618 sekmeler `Note` / `Invitation list`, düğmeler `Copy` | `Copy rows` + `Save to Excel`; alt bant `Source files` (her grup: ad, dosya adı, `Choose…`). Renkler `ozenli-arayuz/references/sistem.md`; koyu/açık. `Pull` ağır işi iş parçacığında yapar, ekran donmaz. `Sonuc.tamam` false ise davet düğmeleri kapalı ve nedeni yazılı.

- [ ] `test_selftest_0`, `test_pull_sonuc_gosterir` (Tk `withdraw`, fixture dosyalarla, sol panelde "Concrete 7d · 2" metni var), `test_hata_varken_kaydet_kapali`.
- [ ] FAIL → yaz → PASS.
- [ ] Göz kontrolü: koyu/açık ekran görüntüsü, 1366×768 ve 1920×1080; sekme geçişi ve Pull sırasında 16 ms sayaçla kare süresi < 40 ms.

### Task 8: Lisans, güncelleme, yayın hattı (CheckList'ten)

**Files:** Copy+rename from `~/Projeler/CheckList`: `launcher.py`, `license_client.py`, `src/updater.py`, `build/publish.py`, `build/CheckList.spec`→`build/CrushCall.spec`, `build/installer.iss`, `build/make_icon.py`, `Guncelleme-Gonder.bat`, `tests/test_launcher.py`, `tests/test_updater_launcher.py`, `tests/test_updater_setup.py`, `tests/test_publish.py`, `tests/test_license_client.py`; `assets/app.ico` (yeni ikon).

Değişecekler: `CheckList`→`CrushCall`, `CHECKLIST_LAUNCHER`→`CRUSHCALL_LAUNCHER`, `checklist_pub_`→`crushcall_pub_`, `GITHUB_OWNER="yavuzeynula"`, `GITHUB_REPO="CrushCall"`, installer `AppId=CrushCall`, `license_client.PROGRAM_ID="crushcall"`, `TITLE="CrushCall"`. Launcher'daki gömülü importlar aynı kalır.

- [ ] Kopyala, `grep -rni checklist` → 0 sonuç (bilinçli kalanlar hariç, gerekçeli).
- [ ] Taşınan testler PASS.

### Task 9: Kabul testi aracı

**Files:** Create `araclar/karsilastir.py`, `tests/test_karsilastir.py`

**Interfaces:** `python araclar/karsilastir.py --kaynaklar <klasör> --davet "Daily Test Invitation TEMPLATE.xlsx" --not <not.txt> --gun 20` — son 20 iş günü için programın çıktısını, şablondaki `(gg.aa.yyyy)` sayfalarının Sample ID ve RIA…Element hücreleriyle ve not dosyasındaki `///…\\\` bloklarıyla karşılaştırır; her fark: tarih, grup, Sample ID, sütun, beklenen, bulunan. Fark yoksa "0 differences" ve çıkış kodu 0.

- [ ] `test_fark_yok`, `test_fark_raporu` (fixture'larla) → FAIL → yaz → PASS.
- [ ] Yavuz'un gerçek dosyaları gelince çalıştır; fark sıfırlanana kadar Task 2–5'e dön.

### Task 10: Panel kaydı, GitHub deposu, ilk yayın

- [ ] LisansPaneli: `programs` tablosuna `crushcall` (LisansPaneli/CLAUDE.md yolu; şifreye dokunma); SECRET → `license_secret.txt`.
- [ ] **Yavuz'a sor:** `yavuzeynula/CrushCall` GitHub deposunu açayım mı (güncelleme için açık depo; `ornekler/` girmez).
- [ ] `python3 ~/tools/yayin/vm.py build/publish.py 0.1.0 "first version" --deneme` → selftest geçer.
- [ ] Windows'ta: kurulum, lisans "Pending" → panelden izin → açılır; `Copy rows` Excel'e yapıştır, örnekle göz karşılaştırması; `Save to Excel` açılır.
- [ ] Yavuz onaylayınca gerçek yayın.
