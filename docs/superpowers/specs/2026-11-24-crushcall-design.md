# CrushCall — Tasarım

Tarih: 2026-11-24 · Durum: onay bekliyor

## Amaç ve neden

Yavuz her gün yarın kırılacak numuneleri elle hazırlıyor:
`OTO-INFO-TAKER` (Power Query) → `Book1` (formüllü ara tablo) → iki çıktı (not metni + davet listesi Excel'i).

Sorunlar:
- Beton kırım dosyası dolup bölününce 12 sorgudaki dosya yolunu tek tek düzeltmek gerekiyor.
- Sorgular hata olunca sessizce boş liste veriyor ("yarın numune yok" sanılabiliyor). `GRT-28days` sorgusunda sütun adı yazım hatası var ("28 day" / "28 days").
- Aynı numunenin satırlarını birleştirmek, seriyi ve dünden devamı kontrol etmek, iki çıktıyı yazmak elle yapılıyor.

CrushCall bu zinciri tek pencereye toplar: dosyaları seç, tarihi seç, **Çek**, kontrol et, iki çıktıyı al.

## Hedefler

1. Her grup için kaynak dosyayı Yavuz seçer; program seçimi hatırlar. Dosya bölününce yalnız o grubun dosyası yeniden seçilir.
2. Seçilen tarihte (varsayılan: yarın, değiştirilebilir) testi olan numuneleri çeker.
3. Aynı Sample ID'nin satırlarını tek satırda birleştirir.
4. Satırları işaretler: SHOTCRETE, NO RIA NUMBER, birleştirildi.
5. Seri kontrolü yapar: gün içinde atlama var mı, önceki test gününün son numarasından devam ediyor mu.
6. Grout'ta 1 günlük / 7 günlük ayrımını döküm tarihinden yapar.
7. İki çıktı üretir: not metni (panoya kopyala) ve davet listesi Excel'i (kaydet).
8. Hiçbir hata gizlenmez: sütun bulunamazsa, tarih okunamazsa, dosya açılamazsa ekranda kırmızı uyarı çıkar (dosya adı, satır, sütun adı ile).

## Hedef olmayanlar

- Kaynak dosyalara yazmak (yalnız okunur).
- `Book1` ya da `Daily Test Invitation TEMPLATE.xlsx` dosyasına yazmak; mail göndermek; yazdırmak.
- Power Query'yi düzeltmek ya da korumak.

## Kesinlikle olmayacaklar

- Hata varken boş ya da eksik listeyi "tamam" gibi göstermek.
- Tahminle doldurmak (ör. grout günü belirsizse 7 yazmak).

## Kısıtlar

- İş bilgisayarı Windows; kaynaklar OneDrive/SharePoint eşitlenmiş klasörde (şirket OneDrive klasörü, `...\Concrete Followup\`). Kod Mac'te de çalışır.
- Yol birleştirme `pathlib`; dosyalar `encoding="utf-8"`; ayar dosyası `Path.home()` altında.
- Kod çok basit ve temiz olacak (Yavuz'un isteği).
- Paketler sürüm + parmak iziyle kilitli (`requirements.in` → `pip-compile --generate-hashes`).

## Yaklaşımlar

| Yol | Artısı | Eksisi |
|---|---|---|
| **A. Python + Tk masaüstü programı (seçilen)** | Bütün adımlar tek tıkla; dosya seçimi hatırlanır; iki çıktı hazır | Exe derlemek gerekir (düzeni zaten var: UTM Windows) |
| B. Excel'de kalıp sorguları düzeltmek | En az iş; yollar tek tabloya toplanır | Birleştirme, seri kontrolü, iki çıktı yine elle |
| C. Tarayıcıda çalışan sayfa | Kurulum yok | Paylaşılan klasördeki dosyaları okumak zahmetli |

Yavuz A'yı seçti (2026-11-24).

## Gruplar ve kaynaklar

Her grup bir kaynak dosyadır. Her dosyada sayfa: `Concrete Comp Log` ya da `Grout Comp Log`. İlk 3 satır atlanır, 4. satır başlıktır. Test tarihi sütunlarının hücresinde **testin yapılacağı tarih** yazar; seçilen tarihe eşit olan satırlar alınır.

| Grup | Kaynak dosya (bugünkü adı) | Yaşlar ve tarih sütunu | Nota | Davete |
|---|---|---|---|---|
| Beton | `SECTION 1-2 Concrete Compression Strength (conformity) (implemented) (updated) from Sheet 04050 -.xlsx` | 7g: `7 days Compressive Strength EN12390-3`, 28g: `28 days Compressive Strength EN12390-3` | evet | evet |
| CMD | `Concrete Compression Strength (conformity) (DESIGNS).xlsx` | 3/7/28g: `N days Compressive Strength EN12390-3` | evet | evet |
| Grout-E (Elastomeric) | `SECTION 1-2 Grout Compression Strength for Elastomeric Bearing and Pillows.xlsx` | 1g: `1 day Compressive Strength EN 196-1`, 28g: başlığı kaynakta doğrulanacak ("28 day…" / "28 days…") | evet | evet |
| Grout-S (Slope/Post-tensioning) | `SECTION 1-2 Grout Compression Strength for Slope Protection Anchors and Post-tensioning.xlsx` | 1-ya-da-7g: `1-day or 7-days Compressive Strength EN 196-1`, 28g: `28-days Compressive Strength EN 196-1` | evet | evet |
| T-CON | `Concrete Compression Strength (conformity) (implemented) - Without RIA.xlsx` | 3/7/28g: `N days Compressive Strength EN12390-3` | evet | hayır |

Alınan alanlar (başlık adına göre bulunur, sıra numarasına göre değil):
`RIA No:`, `Sample ID`, `Mix No`, `Supplier` / `Production Location`, `Pouring Date`, `Sampling Location`, `Location` ve onun sağındaki 3 sütun.
`Location` birleştirilmiş başlıktır: o sütun **Section**, sağındakiler sırayla **Chainage**, **Structure**, **Element**. CMD için ayrıca `Concrete Class`.

Başlık bulunamazsa grup durur ve uyarı çıkar: "Beton: '7 days Compressive Strength EN12390-3' sütunu bulunamadı (dosya: …)".

Tarih hücresi boşsa satır atlanır. Sample ID'si boş satırdaki okunamayan tarih sessiz geçer (5. satır "Test date" alt başlığı, boş şablon satırlarındaki `#REF!`); Sample ID'li satırda `#REF!` gibi formül hatası "formula error … fix it in Excel" uyarısı verir. Doluysa ama tarih değilse (ör. "28/10/2026" metni çözülemiyorsa) uyarı satırı çıkar: "Beton satır 812: tarih okunamadı: 'x'". `dd/mm/yyyy` metni tarih sayılır.

## İşleme kuralları

**Birleştirme.** Aynı grup ve yaşta, aynı Sample ID'li ardışık satırlar tek satıra iner. Her alanda farklı değerler satır sonuyla (`\n`) alt alta yazılır, tekrar edenler bir kez yazılır. Section farklıysa ` & ` ile birleşir (ör. `1 & 2`). Ekranda "1714: 2 satır birleştirildi" bilgisi görünür.

**İşaretler** (Book1'deki O sütunu formülüyle aynı):
- SHOTCRETE: Element'te "shotcrete" ya da "slope" geçiyor (büyük-küçük harf fark etmez).
- NO RIA NUMBER: RIA boş ya da `0` (yalnız RIA'sı olan gruplarda: Beton, Grout).

**Grout 1g / 7g.** Grout-S'nin "1-day or 7-days" sütunu için: test tarihi − döküm tarihi = 1 ise `1d`, 7 ise `7d`. Başka bir sayıysa ya da döküm tarihi boşsa yaş "?" olur ve uyarı çıkar: "GRT-0254: 1 mi 7 mi belli değil, bakılmalı".

**Seri kontrolü** (yalnız Beton 7g/28g; grout numaraları yaşa göre karışık olduğundan yanlış alarm veriyordu — şirket bilgisayarı denemesi 2026-11-24. Sample ID'nin sonundaki sayıya bakılır):
- Gün içi: birleştirmeden sonra numaralar sıralanır; arada eksik varsa uyarı: "Beton 7g: 1716 yok".
- Dünden devam: aynı dosya ve aynı sütunda, seçilen tarihten önceki en yakın test günü bulunur. O günün son numarası + 1, bugünün ilk numarasına eşit değilse uyarı: "Beton 7g: önceki gün (24.11.2026) 1712 ile bitti, bugün 1715 ile başlıyor".
- Bunlar uyarıdır; çıktı yine alınabilir.

## Ekran

Tek pencere (Tk). Yavuz üç aday arasından **C**'yi seçti (2026-11-24): solda kontrol, sağda çıktının önizlemesi.

```
┌ CrushCall ─────────────────────────── Tarih: [‹ 25.11.2026 ›] [Çek] ┐
│ ┌ Kontrol (0,382) ──────┐ ┌ Önizleme (0,618) ────────────────────┐ │
│ │ Beton 7g · 7  ✓ devam │ │ [Not] [Davet listesi]     [Kopyala]  │ │
│ │ Beton 28g · 7 ✓ tam   │ │ ///7days-25.11.2026\\\               │ │
│ │ Grout · 2     ✓       │ │ 1. G26-CON-1713                      │ │
│ │ ⚠ 1714, 1598          │ │ 2. G26-CON-1714 SHOTCRETE            │ │
│ │   birleştirildi       │ │ …                                    │ │
│ └───────────────────────┘ └──────────────────────────────────────┘ │
│                                                 [Kaynak dosyalar]  │
└────────────────────────────────────────────────────────────────────┘
```

- **Programın bütün yazıları İngilizce** (Yavuz 2026-11-24). Düğmeler: `Pull`, `Copy`, `Save to Excel`, `Source files`; sekmeler: `Note`, `Invitation list`; uyarılar da İngilizce (ör. "Concrete 7d: 1716 missing"). Bu yazıdaki Türkçe örnekler yalnız anlatım içindir.
- Önizlemede iki sekme var: `Note` (metin; düğme `Copy`) ve `Invitation list` (tablo; düğme `Save to Excel`). Kaydedilen davet listesi `.xlsx` olarak kaydedilir ve hemen Excel'de açılır (Windows: `os.startfile`, Mac: `open`).
- Sol taraf grupları, sayılarını, seri durumunu ve uyarıları gösterir. Bir gruba tıklanınca önizleme o gruba kayar.

- Tarih alanı: varsayılan yarın; ileri/geri oklar ve elle yazma.
- "Kaynak dosyalar": her grubun bir satırı, yanında "Seç" düğmesi ve seçili dosya adı. Ayar dosyası: `Path.home() / ".crushcall" / "ayarlar.json"`.
- Uyarılar listenin altında kırmızı/turuncu.
- Görsel tasarım uygulama aşamasında `ozenli-arayuz` skill'iyle yapılır (koyu/açık ton, yumuşak geçişler).

## Çıktı 1 — Not metni

`ornekler/not-ornegi.jpg` biçiminde. Sıra: Beton 7g, Beton 28g, sonra madde işaretli satırlar (Grout, CMD, T-CON). Boş grup yazılmaz.

```
///7days-25.11.2026\\\
1. G26-CON-1713
2. G26-CON-1714 SHOTCRETE
…
///28days-25.11.2026\\\
1. G26-CON-1594
…
• G26-GRT-0254 7days
• G26-GRT-0240 28days
• CMD-0123 28days
```

Birleştirilen (DOUBLE) numunenin sonuna birleşen satır sayısı yazılır: `2. G26-CON-1714 SHOTCRETE 2 SET` (3 satırsa `3 SET`). RIA'sız numune notta kalır, yanına `NO RIA` yazılır.

Sample ID kaynaktaki gibi yazılır (CMD'de `CMD-0123`, başına ek yok). "Notu kopyala" panoya koyar (Tk panosu; Windows ve Mac'te çalışır).

## Çıktı 2 — Davet listesi Excel'i

`ornekler/Laboratory_Daily_Notification_Invitation_List_25.11.2026.xlsx` ve `ornekler/davet-cmd-ornegi.jpg` birebir örnektir.

- Dosya adı: `Laboratory_Daily_Notification_Invitation_List_<gg.aa.yyyy>.xlsx`; sayfa adı `<gg.aa.yyyy>`.
- Üst kısım (logolar, "Contractor", başlık) şablondan gelir; veri 13. satırdan başlar.
- Sıra: Beton 7g, Beton 28g, CMD, Grout. `#` 1'den başlar.
- **NO RIA NUMBER olan Beton/Grout numuneleri davet listesine girmez** (yalnız notta görünür). CMD'nin RIA'sı zaten yoktur, `-` ile girer.

| Sütun | Beton | CMD | Grout |
|---|---|---|---|
| Laboratory | BEJV LAB | BEJV LAB | BEJV LAB |
| Test | `Compressive Strength (7d)` | `Compressive Strength (28d)` | `Grout` + satır sonu + `Compressive Strength (7d)` |
| Date | gg.aa.yyyy | aynı | aynı |
| RIA number | RIA | `-` | RIA |
| Test Method | `MKC EN 12390-3:2019` | `-` | `MKC EN 12390-3:2019` (örnek dosyadaki gibi; Yavuz 2026-11-24: "aynı görünmeli") |
| Section | Section | `1` | Section |
| Chainage | Chainage | `11+500` + satır sonu + `QC Lab` | Chainage |
| Structure | Structure | `-` | Structure |
| Element/Location | Element | `Concrete Mix design  #<Mix No> <Concrete Class>` | Element |
| Estimated time | 09:00 | 11:00 | 09:00 |

- Biçim: veri satırları sarı dolgu, kenarlıklı; CMD ve Grout'ta Test hücresi yeşil. Satır yüksekliği satırdaki en çok satır sonu sayısına göre.
- Kaydetme yeri: ilk seferde sorulur, sonra hatırlanır. Kaydettikten sonra dosya Excel'de açılır.
- **Sütunlar ve biçim örnek dosyayla birebir aynı** (11 sütun, aynı sıra, aynı yazı/dolgu/kenarlık, hücre içi satır sonları korunur). Amaç: Yavuz açılan dosyadan satırları kopyalayıp başka Excel'e olduğu gibi yapıştırabilsin. Ekrandaki önizleme de aynı 11 sütunu aynı sırayla gösterir (kısaltma yok).
- **`Copy rows` düğmesi** (Yavuz 2026-11-24): veri satırlarını dosya açmadan panoya koyar; **Excel'e yapıştırınca örnek sarı dosyadaki biçimle gelir** (sarı dolgu, CMD/Grout Test hücresi yeşil, kenarlık, ortalı ve kaydırmalı yazı, hücre içi alt alta satırlar).
  - **Yalnız sarı veri satırları kopyalanır; başlık satırı ve üst kısım panoya girmez** (Yavuz 2026-11-24: "sonradan başlıkları silme işi olmasın"). Yapıştırınca doğrudan diğer listenin altına eklenebilir.
  - Panoya iki biçim birden konur: HTML tablo (biçimli; Excel bunu alır) + sekmeyle ayrılmış düz metin (yedek).
  - Windows: `ctypes` ile "HTML Format" pano kaydı (ek paket yok). Mac: `osascript` ile HTML pano verisi.
  - Hücre içi satır sonu HTML'de `<br style="mso-data-placement:same-cell">` olarak yazılır, yoksa Excel satırı böler.
  - Windows Excel'e yapıştırıp örnek dosyayla göz kontrolü yapılır.

## Doğruluk güvencesi (veri kaybı olmayacak)

Yavuz (2026-11-24): "Veri kaybı olursa işime son verebilirler." Bu bölüm diğer bütün kararların önündedir.

1. **Her satırın hesabı ekranda.** Her grup için: "Kaynakta bu tarihte N satır → çıktıda M numune (K satır birleştirildi, R satır NO RIA olduğu için davetten çıkarıldı)". N = M + K + R tutmazsa kırmızı hata; çıktı düğmeleri kapanır.
2. **Okunamayan hiçbir şey sessiz geçmez.** Tarih sütununda tarih olmayan dolu hücre, formülü olup hesaplanmış değeri olmayan hücre (dosya Excel'de kaydedilmemiş), bulunamayan başlık, açılamayan dosya: hepsi satır numarasıyla listelenir.
3. **Tarih okuma kuralı tek ve açık:** Excel tarih hücresi doğrudan; metinse yalnız `gg/aa/yyyy` ya da `gg.aa.yyyy`. Başka biçim → uyarı. Ay/gün karışması tahmin edilmez.
4. **Kaynağa hiç yazılmaz.** Dosyalar yalnız okuma kipinde açılır; başkası açık tutuyorsa da okunabilir.
5. **Geçmiş günlerle karşılaştırma (kabul testi).** Program, Yavuz'un elle hazırladığı geçmiş çıktılarla karşılaştırılır: `Daily Test Invitation TEMPLATE.xlsx` içindeki günlük sayfalar (her gün ayrı sayfa) ve not dosyasındaki `///7days-…\\\` blokları. En az son 20 iş günü için Sample ID listesi, satır sayısı ve davet satırlarının her hücresi karşılaştırılır. Her fark tek tek açıklanır: programın hatasıysa düzeltilir ve teste eklenir, elle yapılmış bir hataysa Yavuz'a gösterilir. Fark sıfırlanmadan program "hazır" sayılmaz.
6. **İlk hafta iki yol birlikte.** Yavuz bir hafta eski yöntemle de hazırlar, sonuçlar her gün karşılaştırılır.
7. **Otomatik testler** her kuralı küçük örnek dosyalarla sınar: birleştirme, NO RIA, SHOTCRETE, seri atlama, dünden devam, grout 1/7, boş ve bozuk tarih, eksik başlık, aynı tarihte 0 ve çok numune.

## Lisans, güncelleme, yayın (diğer programlar gibi)

Yavuz (2026-11-24): "Güncelleme sistemi ve lisans panelinden izin sistemi olmalı, diğer programlar gibi; herkes kullanamasın." Düzen **CheckList'ten birebir** alınır (kanıtlanmış, aynı ekip):

- `launcher.py` (exe'ye gömülü, nadiren değişir): lisans kapısı → `src/main.py`'yi yükler → yeni kod çökerse `src_backup/`'a döner. `--selftest <klasör>` yayından önce kullanılır.
- `license_client.py`: `LisansPaneli/client/` şablonundan kopya; `PROGRAM_ID = "crushcall"`, `TITLE = "CrushCall"`. SECRET `license_secret.txt`'de (gitignore'da, sohbete basılmaz).
- Panel kaydı: `LisansPaneli/CLAUDE.md`'deki yol (D1 `programs` tablosuna tek satır, panel şifresine dokunulmaz).
- `src/updater.py`: CheckList'inki; GitHub `yavuzeynula/CrushCall` deposunun son release'inden `src.zip` (ya da gerekince kurulum) indirir, SHA256 doğrular.
- `build/publish.py`, `build/CrushCall.spec`, `build/installer.iss`, `build/make_icon.py`, `Guncelleme-Gonder.bat`: CheckList'inkiler, adlar değişir. Yayın: `python3 ~/tools/yayin/vm.py build/publish.py <sürüm> "not"` (UTM Windows'ta derler).
- Testler: CheckList'in `test_launcher`, `test_updater_*`, `test_publish`, `test_license_client` testleri de taşınır.

## Paketleme ve paketler

- Python 3.12, Tk (standart). Paketler CheckList'le aynı sürümler: `openpyxl==3.1.5`, `pillow==13.0.0` (ikon/açılış), `cryptography==50.0.2` (lisans imzası), `pyinstaller==6.19.0`, `pytest==9.0.3`, `pytest-xdist`. Kilit: `pip-compile --generate-hashes`; kurulum `--require-hashes`.
- `.gitattributes` ile satır sonları sabit.

## Kurucuya bırakılanlar (kurucunun kararı)

- Kod düzeni (öneri: `crushcall/kaynak.py` okuma, `islem.py` birleştirme/işaret/seri, `cikti.py` not + Excel, `ekran.py` Tk; testler `tests/`).
- openpyxl şablon açıp kaydedince logoları düşürür. Şablon programın içinde tutulur, logolar `assets/` içinden aynı konumlara yeniden eklenir. Sonuç dosyası örnekle göz kontrolüyle karşılaştırılır.
- Kaynak dosyalar büyük: `read_only=True, data_only=True` ile okunur; açılış süresi ölçülür.
- Testler gerçek kaynak dosyaların başlık düzenini taklit eden küçük örnek dosyalarla yazılır; ayrıca Yavuz'un gönderdiği gerçek dosyalarla bir kez denenir.

## Yavuz'dan beklenenler

1. Beş kaynak dosyanın hepsinden birer kopya (Beton, RIA'sız, DESIGNS, iki Grout). Başlık satırlarını ve "28 day/days" yazımını gerçek dosyada doğrulamak ve karşılaştırma testini yapmak için.
3. Karşılaştırma için geçmiş çıktılar: `Daily Test Invitation TEMPLATE.xlsx` ve not dosyası (Not Defteri'ndeki metin dosyası).
2. Paket onayı: `openpyxl 3.1.5`, `pyinstaller`, `pytest`, `pytest-xdist`.
