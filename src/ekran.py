"""Ana pencere (C düzeni): solda kontrol, sağda Note / Invitation list önizlemesi.

Ağır iş (Excel okuma) iş parçacığında; ekran donmaz. Kare içinde hesap yok:
tablo yerleşimi Pull bitince bir kez kurulur.
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox

import ayar
import cikti_davet
import cikti_not
import islem
import pano
from tanimlar import GRUPLAR

TEMA = {
    "dark": dict(zemin="#081519", kart="#0f2329", yazi="#eef7f8", soluk="#93aeb5", cizgi="#1f3439",
                 vurgu="#52e0bb", hata="#ff8d84", uyari="#f6c35d", ok="#5eeac4",
                 concrete="#e8935a", grout="#b18cf2", cmd="#6fb8dc", tcon="#93aeb5"),
    "light": dict(zemin="#f1f4f7", kart="#ffffff", yazi="#10282f", soluk="#4f6b74", cizgi="#dfe5e8",
                  vurgu="#0b7f6a", hata="#c0271d", uyari="#9a6400", ok="#067a63",
                  concrete="#b8622a", grout="#7a52c4", cmd="#2a7aa3", tcon="#4f6b74"),
}
SUTUNLAR = [("#", 30), ("Laboratory", 74), ("Test", 128), ("Date", 78), ("RIA number", 200),
            ("Test Method", 128), ("Section", 52), ("Chainage", 160), ("Structure", 150),
            ("Element/Location", 260), ("Estimated time", 74)]


def _yarin() -> date:
    return date.today() + timedelta(days=1)


def dosya_ac(yol: Path) -> None:
    """Dosyayı varsayılan programla açar (Windows: os.startfile, Mac: open)."""
    if sys.platform == "win32":
        os.startfile(yol)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(yol)])
    else:
        subprocess.Popen(["xdg-open", str(yol)])


class Pencere:
    def __init__(self, kok: tk.Tk, surum: str = ""):
        self.kok = kok
        self.ayar = ayar.oku()
        self.tarih = _yarin()
        self.sonuc: islem.Sonuc | None = None
        self.sekme = "note"
        self.surum = surum
        aileler = set(tkfont.families(kok))
        self.yazi_ailesi = next((f for f in ("Segoe UI Variable Text", "Segoe UI", "Helvetica Neue") if f in aileler), "TkDefaultFont")
        self.mono = next((f for f in ("Cascadia Mono", "Consolas", "Menlo") if f in aileler), "Courier")
        kok.title("CrushCall")
        kok.minsize(1100, 640)
        self._kur()

    # ---------- yardımcılar ----------
    @property
    def r(self) -> dict:
        return TEMA[self.ayar.get("tema", "dark")]

    def f(self, px: int, kalin: bool = False):
        return (self.yazi_ailesi, -px, "bold" if kalin else "normal")

    def _dugme(self, ust, metin, komut, ana=False):
        r = self.r
        return tk.Button(ust, text=metin, command=komut, font=self.f(13, ana), relief="flat", bd=0,
                         padx=13, pady=5, cursor="hand2",
                         bg=r["vurgu"] if ana else r["kart"], fg=r["zemin"] if ana else r["yazi"],
                         activebackground=r["cizgi"], activeforeground=r["yazi"],
                         disabledforeground=r["soluk"], highlightthickness=1,
                         highlightbackground=r["vurgu"] if ana else r["cizgi"])

    # ---------- düzen ----------
    def _kur(self):
        r = self.r
        for c in self.kok.winfo_children():
            c.destroy()
        self.kok.configure(bg=r["zemin"])
        ust = tk.Frame(self.kok, bg=r["zemin"])
        ust.pack(fill="x", padx=21, pady=(13, 8))
        tk.Label(ust, text="CrushCall", font=self.f(21, True), bg=r["zemin"], fg=r["yazi"]).pack(side="left")
        self.pull = self._dugme(ust, "Pull", self.cek, ana=True)
        self.pull.pack(side="right")
        self._dugme(ust, "›", lambda: self.gun(1)).pack(side="right", padx=(4, 13))
        self.tarih_var = tk.StringVar(value=f"{self.tarih:%d.%m.%Y}")
        giris = tk.Entry(ust, textvariable=self.tarih_var, width=10, justify="center", font=self.f(13),
                         bg=r["kart"], fg=r["yazi"], insertbackground=r["yazi"], relief="flat",
                         highlightthickness=1, highlightbackground=r["cizgi"], highlightcolor=r["vurgu"])
        giris.pack(side="right", ipady=4)
        giris.bind("<Return>", lambda e: self.tarih_yazildi())
        giris.bind("<FocusOut>", lambda e: self.tarih_yazildi())
        self._dugme(ust, "‹", lambda: self.gun(-1)).pack(side="right", padx=4)
        tk.Label(ust, text="Test date", font=self.f(11), bg=r["zemin"], fg=r["soluk"]).pack(side="right", padx=8)

        govde = tk.Frame(self.kok, bg=r["zemin"])
        govde.pack(fill="both", expand=True, padx=21)
        govde.columnconfigure(0, weight=382, uniform="b")
        govde.columnconfigure(1, weight=618, uniform="b")
        govde.rowconfigure(0, weight=1)

        sol = tk.Frame(govde, bg=r["kart"], highlightthickness=1, highlightbackground=r["cizgi"])
        sol.grid(row=0, column=0, sticky="nsew", padx=(0, 13))
        self.kontrol = tk.Text(sol, wrap="word", bd=0, padx=13, pady=13, bg=r["kart"], fg=r["yazi"],
                               font=self.f(13), cursor="arrow", highlightthickness=0)
        self.kontrol.pack(fill="both", expand=True)
        for ad, renk in (("soluk", r["soluk"]), ("hata", r["hata"]), ("uyari", r["uyari"]), ("ok", r["ok"]),
                         ("concrete", r["concrete"]), ("grout", r["grout"]), ("cmd", r["cmd"]), ("tcon", r["tcon"])):
            self.kontrol.tag_configure(ad, foreground=renk)
        for ad in ("concrete", "grout", "cmd", "tcon"):
            self.kontrol.tag_configure(ad, font=self.f(13, True))
        self.kontrol.tag_configure("kucuk", font=self.f(11), spacing3=8)

        sag = tk.Frame(govde, bg=r["kart"], highlightthickness=1, highlightbackground=r["cizgi"])
        sag.grid(row=0, column=1, sticky="nsew")
        cubuk = tk.Frame(sag, bg=r["kart"])
        cubuk.pack(fill="x", padx=13, pady=(13, 8))
        self.b_not = self._dugme(cubuk, "Note", lambda: self.sekme_sec("note"))
        self.b_not.pack(side="left")
        self.b_davet = self._dugme(cubuk, "Invitation list", lambda: self.sekme_sec("invite"))
        self.b_davet.pack(side="left", padx=8)
        self.b_kaydet = self._dugme(cubuk, "Save to Excel", self.kaydet)
        self.b_satir = self._dugme(cubuk, "Copy rows", self.satir_kopyala)
        self.b_kopya = self._dugme(cubuk, "Copy", self.not_kopyala)
        self.neden = tk.Label(cubuk, text="", font=self.f(11), bg=r["kart"], fg=r["hata"])
        self.neden.pack(side="right", padx=8)
        self.icerik = tk.Frame(sag, bg=r["kart"])
        self.icerik.pack(fill="both", expand=True, padx=13, pady=(0, 13))
        self.not_metni = tk.Text(self.icerik, wrap="none", bd=0, bg=r["kart"], fg=r["yazi"],
                                 font=(self.mono, -13), highlightthickness=0, padx=4)
        self.tablo = tk.Canvas(self.icerik, bg=r["kart"], highlightthickness=0)
        self.t_yatay = tk.Scrollbar(self.icerik, orient="horizontal", command=self.tablo.xview)
        self.t_dikey = tk.Scrollbar(self.icerik, orient="vertical", command=self.tablo.yview)
        self.tablo.configure(xscrollcommand=self.t_yatay.set, yscrollcommand=self.t_dikey.set)

        alt = tk.Frame(self.kok, bg=r["zemin"])
        alt.pack(fill="x", padx=21, pady=(8, 13))
        self.durum = tk.Label(alt, text="Press Pull to read the source files.", font=self.f(11),
                              bg=r["zemin"], fg=r["soluk"])
        self.durum.pack(side="left")
        if self.surum:
            tk.Label(alt, text=f"v{self.surum}", font=self.f(11), bg=r["zemin"], fg=r["soluk"]).pack(side="right")
        self._dugme(alt, "Light" if self.ayar.get("tema", "dark") == "dark" else "Dark", self.tema_degis).pack(side="right", padx=8)
        self._dugme(alt, "Source files", self.kaynaklar).pack(side="right")
        self._goster()

    # ---------- tarih ----------
    def gun(self, d: int):
        self.tarih += timedelta(days=d)
        self.tarih_var.set(f"{self.tarih:%d.%m.%Y}")
        self._temizle("Date changed. Press Pull.")

    def tarih_yazildi(self):
        try:
            yeni = datetime.strptime(self.tarih_var.get().strip(), "%d.%m.%Y").date()
        except ValueError:
            self.tarih_var.set(f"{self.tarih:%d.%m.%Y}")
            self.durum.configure(text="Date must be dd.mm.yyyy", fg=self.r["hata"])
            return
        if yeni != self.tarih:
            self.tarih = yeni
            self._temizle("Date changed. Press Pull.")

    def _temizle(self, mesaj: str):
        self.sonuc = None
        self.durum.configure(text=mesaj, fg=self.r["soluk"])
        self._goster()

    # ---------- çekme ----------
    def cek(self):
        self.tarih_yazildi()
        secim = self.ayar.get("kaynaklar", {})
        self.pull.configure(text="Reading…", state="disabled")
        kutu = {}

        def is_():
            try:
                kutu["sonuc"] = islem.hesapla(self.tarih, secim)
            except Exception as e:  # beklenmeyen: ekranda göster, gizleme
                kutu["hata"] = e
        threading.Thread(target=is_, daemon=True).start()

        def bekle():
            if "sonuc" not in kutu and "hata" not in kutu:
                self.kok.after(50, bekle)
                return
            self.pull.configure(text="Pull", state="normal")
            if "hata" in kutu:
                messagebox.showerror("CrushCall", f"Pull failed:\n\n{kutu['hata']}")
                return
            self.sonuc = kutu["sonuc"]
            self._goster()
        bekle()

    # ---------- gösterim ----------
    def _goster(self):
        self._kontrol_yaz()
        self.sekme_sec(self.sekme, gecis=False)

    def _kontrol_yaz(self):
        k = self.kontrol
        k.configure(state="normal")
        k.delete("1.0", "end")
        s = self.sonuc
        if s is None:
            k.insert("end", "Nothing pulled yet.", "soluk")
        else:
            for h in s.hesap:
                if h.kaynak_satir == 0:
                    continue
                kod = h.ad.split()[0].lower().replace("-", "")
                renk = "grout" if kod.startswith("grout") else ("tcon" if kod == "tcon" else kod)
                k.insert("end", f"{h.ad} · {h.numune}\n", renk)
                ek = f"{h.kaynak_satir} source rows → {h.numune} samples"
                ek += f" · {h.birlesen} merged" if h.birlesen else ""
                ek += f" · {h.no_ria} no RIA" if h.no_ria else ""
                ek += " · note only" if h.sadece_not else ""
                k.insert("end", ek + "\n", ("soluk", "kucuk"))
            if not any(h.kaynak_satir for h in s.hesap):
                k.insert("end", f"No samples for {self.tarih:%d.%m.%Y}.\n", "soluk")
            for u in s.uyarilar:
                k.insert("end", ("✕ " if u.seviye == "error" else "⚠ ") + u.metin + "\n",
                         ("hata" if u.seviye == "error" else "uyari", "kucuk"))
            for n in s.numuneler:
                if n.set_sayisi > 1:
                    k.insert("end", f"⚠ {n.sample_id}: {n.set_sayisi} rows merged\n", ("uyari", "kucuk"))
                if n.no_ria:
                    k.insert("end", f"⚠ {n.sample_id}: NO RIA, left out of invitation\n", ("uyari", "kucuk"))
            toplam = len(s.numuneler)
            self.durum.configure(text=f"{toplam} samples · every source row accounted for" if s.tamam
                                 else "Errors found: fix them before using the outputs",
                                 fg=self.r["soluk"] if s.tamam else self.r["hata"])
        k.configure(state="disabled")

    def sekme_sec(self, sekme: str, gecis: bool = True):
        self.sekme = sekme
        r = self.r
        for b, ad in ((self.b_not, "note"), (self.b_davet, "invite")):
            b.configure(highlightbackground=r["vurgu"] if ad == sekme else r["cizgi"],
                        fg=r["vurgu"] if ad == sekme else r["yazi"])
        for w in (self.b_kaydet, self.b_satir, self.b_kopya, self.not_metni, self.tablo, self.t_yatay, self.t_dikey):
            w.pack_forget()
            w.place_forget()
        tamam = self.sonuc is not None and self.sonuc.tamam
        if sekme == "note":
            self.b_kopya.pack(side="right")
            self.b_kopya.configure(state="normal" if self.sonuc else "disabled")
            self._not_doldur()
            hedef = self.not_metni
        else:
            self.b_kaydet.pack(side="right")
            self.b_satir.pack(side="right", padx=8)
            for b in (self.b_kaydet, self.b_satir):
                b.configure(state="normal" if tamam else "disabled")
            self._tablo_ciz()
            hedef = self.tablo
        self.neden.configure(text="" if tamam or self.sonuc is None else "Outputs locked: see errors on the left")
        if sekme == "invite":
            self.t_yatay.pack(side="bottom", fill="x")
            self.t_dikey.pack(side="right", fill="y")
        hedef.pack(fill="both", expand=True)
        if gecis:
            self._kaydir(hedef)

    def _kaydir(self, w, adim=0, adimlar=11):
        """Yumuşak geçiş: içerik 13 px aşağıdan yerine oturur (yalnız konum; ~180 ms)."""
        if adim == 0:
            w.pack_forget()
            w.place(x=0, y=13, relwidth=1, relheight=1)
        t = adim / adimlar
        egri = 1 - (1 - t) ** 3
        if adim >= adimlar:
            w.place_forget()
            w.pack(fill="both", expand=True)
            return
        w.place_configure(y=round(13 * (1 - egri)))
        self.kok.after(16, self._kaydir, w, adim + 1, adimlar)

    def _not_doldur(self):
        t = self.not_metni
        t.configure(state="normal")
        t.delete("1.0", "end")
        t.insert("end", cikti_not.yap(self.sonuc.numuneler, self.tarih) if self.sonuc else "Nothing pulled yet.")
        t.configure(state="disabled")

    def _davet_satirlari(self):
        return cikti_davet.satirlar(self.sonuc.numuneler, self.tarih) if self.sonuc else []

    def _tablo_ciz(self):
        """Excel görünüşünde tablo: sarı satırlar, siyah kenar, grout/CMD Test hücresi yeşil."""
        c = self.tablo
        c.delete("all")
        satirlar = self._davet_satirlari()
        if not satirlar:
            c.create_text(8, 8, anchor="nw", text="Nothing pulled yet." if not self.sonuc else "No rows for the invitation.",
                          fill=self.r["soluk"], font=self.f(13))
            c.configure(scrollregion=(0, 0, 10, 10))
            return
        yazi = ("Calibri", -13)
        x0, y = 0, 0
        baslik = [ad for ad, _ in SUTUNLAR]
        for no, satir in enumerate([baslik] + satirlar):
            yukseklik = 0
            kimlikler = []
            x = x0
            for k, (_, gen) in enumerate(SUTUNLAR):
                tid = c.create_text(x + gen / 2, y + 4, anchor="n", width=gen - 8, justify="center",
                                    text=str(satir[k]).replace("\r\n", "\n"), fill="#000000",
                                    font=(yazi[0], yazi[1], "bold") if no == 0 else yazi)
                kimlikler.append((tid, x, gen))
                b = c.bbox(tid)
                yukseklik = max(yukseklik, b[3] - b[1] + 8, 30)
                x += gen
            for k, (tid, x, gen) in enumerate(kimlikler):
                renk = "#ffffff" if no == 0 or k == 0 else ("#92D050" if k == 2 and cikti_davet._yesil_mi(satir) else "#FFFF00")
                rid = c.create_rectangle(x, y, x + gen, y + yukseklik, fill=renk, outline="#000000")
                c.tag_lower(rid, tid)
                b = c.bbox(tid)
                c.move(tid, 0, (yukseklik - (b[3] - b[1])) / 2 - 4)
            y += yukseklik
        c.configure(scrollregion=(0, 0, x, y))

    # ---------- çıktılar ----------
    def not_kopyala(self):
        if not self.sonuc:
            return
        self.kok.clipboard_clear()
        self.kok.clipboard_append(cikti_not.yap(self.sonuc.numuneler, self.tarih))
        self._onay(self.b_kopya, "Copied")

    def satir_kopyala(self):
        """Windows: dosyayı Excel'de açıp satırları Excel'e kopyalatır ("Insert Copied Cells" çıksın).
        Olmazsa (ya da Mac'te): biçimli HTML panoya."""
        s = self._davet_satirlari()
        if sys.platform == "win32":
            try:
                yol = cikti_davet.kopya_dosyasi(s, self.tarih)
            except (cikti_davet.KayitHatasi, OSError) as e:
                messagebox.showerror("CrushCall", str(e))
                return
            self.b_satir.configure(text="Opening Excel…", state="disabled")
            kutu = {}
            threading.Thread(target=lambda: kutu.setdefault("ok", pano.excel_kopyala(yol, len(s))), daemon=True).start()

            def bekle():
                if "ok" not in kutu:
                    self.kok.after(100, bekle)
                    return
                self.b_satir.configure(text="Copy rows", state="normal")
                if kutu["ok"]:
                    self.durum.configure(text="Rows copied in Excel. In your file: right-click a row number › Insert Copied Cells",
                                         fg=self.r["ok"])
                else:
                    self._html_kopyala(s)
            bekle()
            return
        self._html_kopyala(s)

    def _html_kopyala(self, s):
        bicimli = pano.kopyala(cikti_davet.html(s), cikti_davet.tsv(s), self.kok)
        self._onay(self.b_satir, f"{len(s)} rows copied" if bicimli else f"{len(s)} rows copied (plain)")

    def _dosyaya_yaz(self) -> Path | None:
        """Davet dosyasını kayıt klasörüne yazar (klasör ilk sefer sorulur). Hata ekranda."""
        klasor = self.ayar.get("kayit_klasoru")
        if not klasor or not Path(klasor).is_dir():
            klasor = filedialog.askdirectory(title="Where should invitation lists be saved?")
            if not klasor:
                return None
            self.ayar["kayit_klasoru"] = klasor
            ayar.yaz(self.ayar)
        try:
            return cikti_davet.kaydet(self._davet_satirlari(), self.tarih, Path(klasor))
        except cikti_davet.KayitHatasi as e:
            messagebox.showerror("CrushCall", str(e))
            return None

    def kaydet(self):
        yol = self._dosyaya_yaz()
        if yol:
            self._onay(self.b_kaydet, "Saved, opening…")
            dosya_ac(yol)

    def _onay(self, dugme, metin):
        eski = dugme.cget("text")
        dugme.configure(text=metin)
        self.kok.after(1400, lambda: dugme.winfo_exists() and dugme.configure(text=eski))

    # ---------- ayarlar ----------
    def tema_degis(self):
        self.ayar["tema"] = "light" if self.ayar.get("tema", "dark") == "dark" else "dark"
        ayar.yaz(self.ayar)
        self._kur()

    def kaynaklar(self):
        r = self.r
        p = tk.Toplevel(self.kok, bg=r["zemin"], padx=21, pady=21)
        p.title("Source files")
        p.transient(self.kok)
        tk.Label(p, text="Choose the log file for each group", font=self.f(17, True),
                 bg=r["zemin"], fg=r["yazi"]).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 13))
        secim = self.ayar.setdefault("kaynaklar", {})
        for i, g in enumerate(GRUPLAR, 1):
            tk.Label(p, text=g.ad, font=self.f(13, True), bg=r["zemin"], fg=r["yazi"]).grid(row=i, column=0, sticky="w", pady=5)
            ad = tk.Label(p, text=Path(secim[g.kod]).name if secim.get(g.kod) else "not chosen", font=self.f(11),
                          bg=r["zemin"], fg=r["soluk"] if secim.get(g.kod) else r["hata"], width=60, anchor="w")
            ad.grid(row=i, column=1, sticky="w", padx=13)

            def sec(g=g, ad=ad):
                yol = filedialog.askopenfilename(parent=p, title=f"{g.ad} log file",
                                                 filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
                if yol:
                    secim[g.kod] = yol
                    ayar.yaz(self.ayar)
                    ad.configure(text=Path(yol).name, fg=r["soluk"])
                    self._temizle("Source changed. Press Pull.")
            self._dugme(p, "Choose…", sec).grid(row=i, column=2)
