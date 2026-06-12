"use client";

import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { findMatches, kisaltmaAc, normalize } from "@/lib/fuzzy";
import type { StokItem, ScannedItem, ConfirmedItem, ScanResult } from "@/lib/types";
import bundledStok from "@/lib/stokData.json";

type Step = "ral" | "scan" | "processing" | "confirm" | "more_pages" | "done" | "update";

// ── SVG Icons ─────────────────────────────────────────────────────────────
function IconCamera() {
  return (
    <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
      <path d="M23 19a2 2 0 01-2 2H3a2 2 0 01-2-2V8a2 2 0 012-2h4l2-3h6l2 3h4a2 2 0 012 2z"/>
      <circle cx="12" cy="13" r="4"/>
    </svg>
  );
}
function IconGallery() {
  return (
    <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="3" width="18" height="18" rx="3"/>
      <circle cx="8.5" cy="8.5" r="1.5"/>
      <polyline points="21 15 16 10 5 21"/>
    </svg>
  );
}
function IconChevron() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="9 18 15 12 9 6"/>
    </svg>
  );
}
function IconRefresh() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="23 4 23 10 17 10"/>
      <path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/>
    </svg>
  );
}

// ── Helpers ────────────────────────────────────────────────────────────────
/** Parst Mengenangaben wie "1848 + 1900" → 3748, oder "280" → 280 */
function parseMiktar(raw: string | number | undefined): string {
  if (raw === undefined || raw === null) return "";
  const s = String(raw);
  // Enthält "+"? → alle Zahlen addieren
  if (s.includes("+")) {
    const sum = s
      .split("+")
      .map(p => parseFloat(p.replace(/[^\d.,]/g, "").replace(",", ".")) || 0)
      .reduce((a, b) => a + b, 0);
    return sum % 1 === 0 ? String(sum) : sum.toFixed(2);
  }
  return s.trim();
}

function toBase64(file: File): Promise<string> {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.readAsDataURL(file);
    r.onload  = () => res((r.result as string).split(",")[1]);
    r.onerror = rej;
  });
}
function bugunStr() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}.${d.getFullYear()}`;
}
const OGRENIM_KEY = "ogrenilen_eslesmeler";

/** Öğrenilen düzeltmeyi kaydet: "ATLAS HBSB" → stok_kodu */
function ogrenimKaydet(urunAdi: string, stokKodu: string) {
  try {
    const key = normalize(urunAdi);
    const map: Record<string, string> = JSON.parse(
      localStorage.getItem(OGRENIM_KEY) || "{}"
    );
    map[key] = stokKodu;
    localStorage.setItem(OGRENIM_KEY, JSON.stringify(map));
  } catch { /* ignore */ }
}

/** Öğrenilmiş düzeltmeyi getir — varsa stok_kodu döner */
function ogrenimGetir(urunAdi: string): string | null {
  try {
    const key = normalize(urunAdi);
    const map: Record<string, string> = JSON.parse(
      localStorage.getItem(OGRENIM_KEY) || "{}"
    );
    return map[key] || null;
  } catch { return null; }
}

function stokYukle(): StokItem[] {
  const base = bundledStok as StokItem[];
  try {
    const raw = localStorage.getItem("stok_guncellemeler");
    if (!raw) return base;
    const extra: StokItem[] = JSON.parse(raw);
    const map = new Map(base.map(i => [i.stok_kodu, i]));
    for (const e of extra) map.set(e.stok_kodu, e);
    return Array.from(map.values());
  } catch { return base; }
}

// ── Main ───────────────────────────────────────────────────────────────────
export default function Home() {
  const [adim, setAdim]                       = useState<Step>("ral");
  const [adimGecmisi, setAdimGecmisi]         = useState<Step[]>([]);
  const [stokData, setStokData]               = useState<StokItem[]>([]);
  const [tumOnaylananlar, setTumOnaylananlar]  = useState<ConfirmedItem[]>([]);
  const [isEmriNo]                            = useState("");
  const [tarih]                               = useState(bugunStr());
  const [taramaKonusu, setTaramaKonusu]       = useState<ScanResult | null>(null);
  const [mevcutIndex, setMevcutIndex]         = useState(0);
  const [sayfaOnaylananlar, setSayfaOnaylananlar] = useState<ConfirmedItem[]>([]);
  const [onizlemeGorsel, setOnizlemeGorsel]   = useState<string | null>(null);
  const [taramaFotolari, setTaramaFotolari]   = useState<{data: string; mime: string}[]>([]);
  const [hata, setHata]                       = useState<string | null>(null);
  const [ralRenk, setRalRenk]                 = useState("");
  const [odDurum, setOdDurum]                 = useState<"idle"|"loading"|"success"|"error">("idle");
  const [odMesaj, setOdMesaj]                 = useState<string | null>(null);
  const [duzenMiktar, setDuzenMiktar]         = useState("");
  const [duzenStok, setDuzenStok]             = useState<StokItem | null>(null);
  const [stokArama, setStokArama]             = useState("");
  const [stokOneriler, setStokOneriler]       = useState<StokItem[]>([]);
  const [duzenlemeAcik, setDuzenlemeAcik]     = useState(false);
  const [guncellemeMetni, setGuncellemeMetni] = useState<string | null>(null);

  useEffect(() => { setStokData(stokYukle()); }, []);

  /** Adım değiştir ve geçmişe kaydet */
  function gidim(hedef: Step) {
    setAdimGecmisi(prev => [...prev, adim]);
    setAdim(hedef);
  }

  /** Bir önceki adıma dön */
  function geriGit() {
    setAdimGecmisi(prev => {
      if (prev.length === 0) return prev;
      const yeni = [...prev];
      const onceki = yeni.pop()!;
      setAdim(onceki);
      // processing'e geri dönmeyelim
      return yeni;
    });
  }

  const geriVar = adimGecmisi.length > 0 && adim !== "ral" && adim !== "done";

  async function gorselSec(e: React.ChangeEvent<HTMLInputElement>) {
    const dosya = e.target.files?.[0];
    if (!dosya) return;
    setHata(null);
    setOnizlemeGorsel(URL.createObjectURL(dosya));
    gidim("processing");
    try {
      const b64 = await toBase64(dosya);
      // Fotoğrafı sakla (OneDrive yüklemesi için)
      setTaramaFotolari(prev => [...prev, { data: b64, mime: dosya.type || "image/jpeg" }]);
      const res = await fetch("/api/scan", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ imageBase64: b64, mimeType: dosya.type }),
      });
      if (!res.ok) throw new Error((await res.json()).error || "Tarama başarısız");
      const sonuc: ScanResult = await res.json();
      if (!sonuc.items?.length) throw new Error("Hiçbir ürün algılanamadı.");
      // Sadece manuel RAL girilmemişse AI'dan okunanı kullan
      if (!ralRenk && sonuc.ral_renk) setRalRenk(sonuc.ral_renk);
      // Miktar içinde "+" varsa topla (örn: "1848 + 1900" → "3748")
      sonuc.items = sonuc.items.map(item => ({
        ...item,
        miktar: parseMiktar(item.miktar),
      }));
      setTaramaKonusu(sonuc); setMevcutIndex(0); setSayfaOnaylananlar([]); gidim("confirm");
    } catch (err) { setHata(String(err)); setAdim("scan"); }
  }

  const satirHazirla = useCallback((urun: ScannedItem) => {
    setDuzenMiktar(urun.miktar || ""); setStokArama(urun.urun_adi || ""); setDuzenlemeAcik(false);

    // ── 1. Öğrenilmiş düzeltme varsa → direkt kullan ──────────────────────
    const ogrenilmis = ogrenimGetir(urun.urun_adi);
    if (ogrenilmis) {
      const stokItem = stokData.find(s => s.stok_kodu === ogrenilmis);
      if (stokItem) {
        // Öğrenileni başa koy, geri kalanını fuzzy ile doldur
        const diger = findMatches(urun.urun_adi, stokData, 6)
          .filter(s => s.stok_kodu !== ogrenilmis);
        setStokOneriler([stokItem, ...diger]);
        setDuzenStok(stokItem);
        return;
      }
    }

    // ── 2. Bilinen takma adlar → stok arama terimi ────────────────────────
    const ALIAS_MAP: { pattern: RegExp; query: string }[] = [
      { pattern: /sa[çc]\s*kapak/i, query: "SC" },
    ];
    for (const { pattern, query } of ALIAS_MAP) {
      if (pattern.test(urun.urun_adi)) {
        const q = ralRenk ? `${ralRenk} ${query}` : query;
        const e = findMatches(q, stokData, 5);
        setStokOneriler(e); setDuzenStok(e[0] || null);
        return;
      }
    }

    // ── 3. Kısaltma açma: HBSB → Hebe Schiebe vb. ────────────────────────
    // "mat" sadece yüzey bilgisidir, ürün tipini belirtmez → aramadan çıkar
    const temizAd = kisaltmaAc(urun.urun_adi)
      .replace(/\bmat\b/gi, "").replace(/\s+/g, " ").trim();

    // Eğer ürün adında tip belirten bir kelime yoksa → pencere/kapı kolu varsay
    const TIP_KELIMELERI = [
      "gobek","göbek","kapak","mentese","menteş","menteşe",
      "ayna","mandal","kilit","flanş","flans","kose","köşe",
      "cita","çıta","panel","kapı kolu","pencere kolu",
      "kol","kavrama","topuz","pres","sürgü","surgü"
    ];
    const tipVar = TIP_KELIMELERI.some(k =>
      urun.urun_adi.toLowerCase().includes(k)
    );
    const tipEki = tipVar ? "" : " kol";

    const q = ralRenk
      ? `${ralRenk} ${temizAd}${tipEki}`
      : `${temizAd}${tipEki}`;

    const e = findMatches(q, stokData, 5);
    setStokOneriler(e); setDuzenStok(e[0] || null);
  }, [stokData, ralRenk]);

  useEffect(() => {
    if (adim === "confirm" && taramaKonusu && mevcutIndex < taramaKonusu.items.length)
      satirHazirla(taramaKonusu.items[mevcutIndex]);
  }, [mevcutIndex, adim, taramaKonusu, satirHazirla]);

  function onayla() {
    const u = taramaKonusu!.items[mevcutIndex];
    // Onaylanan eşleşmeyi öğren — bir sonraki aynı ürün adında ilk sıraya gelir
    if (duzenStok) ogrenimKaydet(u.urun_adi, duzenStok.stok_kodu);
    ilerle([...sayfaOnaylananlar, { original_urun_adi: u.urun_adi, original_miktar: u.miktar,
      confirmed_stok: duzenStok, confirmed_miktar: parseFloat(duzenMiktar.replace(",",".")) || 0, skipped: false }]);
  }
  function atla() {
    const u = taramaKonusu!.items[mevcutIndex];
    ilerle([...sayfaOnaylananlar, { original_urun_adi: u.urun_adi, original_miktar: u.miktar,
      confirmed_stok: null, confirmed_miktar: 0, skipped: true }]);
  }
  function ilerle(g: ConfirmedItem[]) {
    setSayfaOnaylananlar(g);
    if (mevcutIndex + 1 < taramaKonusu!.items.length) setMevcutIndex(mevcutIndex + 1);
    else { setTumOnaylananlar(prev => [...prev, ...g]); gidim("more_pages"); }
  }
  function stokAramaGuncelle(q: string) {
    setStokArama(q);
    setStokOneriler(q.length >= 1 ? findMatches(q, stokData, 50) : []);
  }
  async function stokGuncelle(e: React.ChangeEvent<HTMLInputElement>) {
    const dosya = e.target.files?.[0]; if (!dosya) return;
    setGuncellemeMetni("Okunuyor...");
    try {
      const buf = await dosya.arrayBuffer();
      const wb  = XLSX.read(buf, { type: "array" });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string,string>>(ws, { raw: false, defval: "" });
      const yeni: StokItem[] = rows.flatMap(r => {
        const kodu   = (r["STOK_KODU"]||r["Stok Kodu"]||"").trim();
        const adi    = (r["STOK_ADI"] ||r["Stok Adı"] ||"").trim();
        const cesit  = (r["Çeşit"]||"").trim();
        const uretim = (r["Üretim"]||r["URETIM"]||r["Uretim"]||"").trim().toUpperCase();
        // Sadece YARI MAMUL
        if (!kodu || !adi) return [];
        if (uretim && uretim !== "YARI MAMUL") return [];
        return [{ stok_kodu: kodu, stok_adi: adi, cesit }];
      });
      const mevcut = new Set((bundledStok as StokItem[]).map(i => i.stok_kodu));
      const gercekYeni = yeni.filter(i => !mevcut.has(i.stok_kodu));
      const eski: StokItem[] = JSON.parse(localStorage.getItem("stok_guncellemeler") || "[]");
      const map = new Map(eski.map(i => [i.stok_kodu, i]));
      for (const y of gercekYeni) map.set(y.stok_kodu, y);
      localStorage.setItem("stok_guncellemeler", JSON.stringify(Array.from(map.values())));
      setStokData(stokYukle());
      setGuncellemeMetni(`✅ ${gercekYeni.length} yeni ürün eklendi.`);
    } catch (err) { setGuncellemeMetni("❌ " + String(err)); }
  }
  async function excelIndir() {
    setHata(null);
    try {
      const res = await fetch("/api/export", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ is_emri_no: isEmriNo, tarih, items: tumOnaylananlar }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href = url; a.download = `Uretim_${tarih.replace(/\./g,"-")}.xlsx`; a.click();
      URL.revokeObjectURL(url);
    } catch (err) { setHata("Dışa aktarma hatası: " + String(err)); }
  }
  function yenidenBaslat() {
    setTumOnaylananlar([]); setTaramaKonusu(null); setSayfaOnaylananlar([]);
    setMevcutIndex(0); setOnizlemeGorsel(null); setRalRenk("");
    setTaramaFotolari([]); setOdDurum("idle"); setOdMesaj(null);
    setAdim("ral");
  }

  async function oneDriveYukle() {
    setOdDurum("loading"); setOdMesaj(null);
    try {
      const { uploadToOneDrive } = await import("@/lib/onedrive");
      const baseName = tarih.replace(/\./g, "-") + "_" +
        new Date().toTimeString().slice(0, 5).replace(":", "-");
      const result = await uploadToOneDrive({
        exportPayload: { is_emri_no: isEmriNo, tarih, items: tumOnaylananlar },
        baseName,
        photos: taramaFotolari,
      });
      setOdDurum("success");
      setOdMesaj(`✅ ${result.uploaded} dosya ${result.folder} klasörüne yüklendi`);
    } catch (err) {
      setOdDurum("error");
      setOdMesaj("❌ " + String(err));
    }
  }

  const mevcutUrun      = taramaKonusu?.items[mevcutIndex];
  const toplamUrun      = taramaKonusu?.items.length || 0;
  const onaylananSayisi = tumOnaylananlar.filter(i => !i.skipped).length;
  const wfBlue          = "#2B5597";

  // ── Render ───────────────────────────────────────────────────────────────
  return (
    <div className="space-y-4 pb-12">

      {/* Geri butonu */}
      {geriVar && (
        <button
          onClick={geriGit}
          className="flex items-center gap-1.5 text-gray-500 font-medium py-1 active:opacity-60 transition-opacity">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="15 18 9 12 15 6"/>
          </svg>
          <span className="text-sm">Geri</span>
        </button>
      )}

      {/* Hata */}
      {hata && (
        <div className="bg-red-50 border border-red-200 text-red-800 rounded-2xl p-4 flex gap-3 items-start">
          <span className="text-lg mt-0.5">⚠️</span>
          <p className="flex-1 text-sm">{hata}</p>
          <button onClick={() => setHata(null)} className="text-red-300 font-bold text-lg leading-none">✕</button>
        </div>
      )}

      {/* ── RAL GİRİŞİ ─────────────────────────────────────────────────── */}
      {adim === "ral" && (
        <div className="wf-fade-up space-y-4">
          <div className="bg-white rounded-3xl shadow-md overflow-hidden">
            <div className="h-1.5 w-full" style={{ background: wfBlue }} />
            <div className="px-6 pt-7 pb-8 space-y-6">

              {/* Başlık */}
              <div>
                <p className="text-[11px] font-bold uppercase tracking-widest mb-1" style={{ color: wfBlue }}>
                  Adım 1 / 2
                </p>
                <h2 className="text-2xl font-black text-gray-900">RAL Renk Kodu</h2>
                <p className="text-gray-400 text-sm mt-1">
                  Fişteki RAL kodunu girin, sonra fotoğraf çekin
                </p>
              </div>

              {/* Büyük RAL input */}
              <div>
                <div className="relative">
                  <span className="absolute left-5 top-1/2 -translate-y-1/2 text-sm font-bold tracking-widest"
                    style={{ color: wfBlue }}>RAL</span>
                  <input
                    type="text"
                    inputMode="text"
                    value={ralRenk}
                    onChange={e => setRalRenk(e.target.value.slice(0, 12))}
                    placeholder="örn: 9005 mat"
                    autoFocus
                    autoCapitalize="none"
                    className="w-full border-2 rounded-2xl pl-16 pr-5 py-5 text-3xl font-black tracking-widest focus:outline-none transition-colors"
                    style={{ borderColor: ralRenk ? wfBlue : "#e5e7eb",
                             color: ralRenk ? wfBlue : "#9ca3af" }}
                  />
                </div>

                {/* Yaygın RAL hızlı seç */}
                <div className="mt-3">
                  <p className="text-xs text-gray-400 font-medium mb-2">Hızlı seç:</p>
                  <div className="flex flex-wrap gap-2">
                    {["9005 mat","8019","7016","7039 mat","9016","1013","6005"].map(ral => (
                      <button key={ral}
                        onClick={() => setRalRenk(ral)}
                        className={`px-3 py-1.5 rounded-xl text-sm font-bold border-2 transition-all ${
                          ralRenk === ral
                            ? "text-white border-transparent"
                            : "bg-white border-gray-200 text-gray-600"
                        }`}
                        style={ralRenk === ral ? { background: wfBlue, borderColor: wfBlue } : {}}>
                        {ral}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Devam butonu */}
              <button
                onClick={() => setAdim("scan")}
                className="w-full text-white font-bold py-5 rounded-3xl text-lg shadow-lg active:scale-95 transition-transform"
                style={{ background: ralRenk
                  ? `linear-gradient(145deg, ${wfBlue}, #1e3d6e)`
                  : "#9ca3af" }}>
                {ralRenk ? `RAL ${ralRenk} ile Devam →` : "RAL girmeden Devam →"}
              </button>

            </div>
          </div>

          {/* Güncelle linki */}
          <button onClick={() => { setAdim("update"); setGuncellemeMetni(null); }}
            className="w-full bg-white border border-gray-200 rounded-2xl py-3.5 flex items-center justify-between px-5 shadow-sm active:bg-gray-50">
            <div className="flex items-center gap-2 text-gray-500">
              <IconRefresh />
              <span className="text-sm font-medium">Stok Kartını Güncelle</span>
            </div>
            <IconChevron />
          </button>
        </div>
      )}

      {/* ── TARAMA ──────────────────────────────────────────────────────── */}
      {adim === "scan" && (
        <div className="space-y-3 wf-fade-up">

          {/* Durum kartı — RAL göster */}
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 px-5 py-4 flex justify-between items-center">
            <div>
              <p className="text-[11px] font-bold uppercase tracking-widest mb-0.5" style={{ color: wfBlue }}>
                Adım 2 / 2 — Fotoğraf
              </p>
              {tumOnaylananlar.length > 0 && (
                <p className="text-sm font-medium" style={{ color: "#16a34a" }}>
                  ✓ {onaylananSayisi} ürün onaylandı
                </p>
              )}
            </div>
            {/* RAL badge */}
            <div className="flex items-center gap-2">
              {ralRenk ? (
                <div className="flex items-center gap-1.5">
                  <span className="text-[11px] font-bold text-white px-3 py-1.5 rounded-xl"
                    style={{ background: wfBlue }}>
                    RAL {ralRenk}
                  </span>
                  <button onClick={() => setAdim("ral")}
                    className="text-gray-400 text-xs underline">değiştir</button>
                </div>
              ) : (
                <button onClick={() => setAdim("ral")}
                  className="text-xs font-bold px-3 py-1.5 rounded-xl border-2 border-dashed border-gray-300 text-gray-400">
                  + RAL ekle
                </button>
              )}
            </div>
          </div>

          {/* ── Kamera & Galeri — eşit boyutlu, premium ── */}
          <div className="grid grid-cols-2 gap-3">
            {/* Kamera */}
            <label className="block cursor-pointer btn-press">
              <div className="bg-white border border-gray-100 rounded-3xl shadow-sm overflow-hidden
                              active:shadow-md active:border-blue-200 transition-all h-full">
                {/* Renkli üst şerit */}
                <div className="h-1.5 w-full" style={{ background: wfBlue }} />
                <div className="flex flex-col items-center justify-center gap-3 py-8 px-4">
                  <div className="w-16 h-16 rounded-2xl flex items-center justify-center shadow-sm"
                    style={{ background: `linear-gradient(145deg, #2B5597, #1e3d6e)` }}>
                    <span style={{ color: "white" }}><IconCamera /></span>
                  </div>
                  <div className="text-center">
                    <p className="font-bold text-gray-900 text-base">Kamera</p>
                    <p className="text-xs text-gray-400 mt-0.5">Fotoğraf çek</p>
                  </div>
                </div>
              </div>
              <input type="file" accept="image/*" capture="environment" className="hidden" onChange={gorselSec} />
            </label>

            {/* Galeri */}
            <label className="block cursor-pointer btn-press">
              <div className="bg-white border border-gray-100 rounded-3xl shadow-sm overflow-hidden
                              active:shadow-md active:border-blue-200 transition-all h-full">
                {/* Renkli üst şerit */}
                <div className="h-1.5 w-full" style={{ background: "#A7A9AC" }} />
                <div className="flex flex-col items-center justify-center gap-3 py-8 px-4">
                  <div className="w-16 h-16 rounded-2xl flex items-center justify-center shadow-sm"
                    style={{ background: "linear-gradient(145deg, #6b7280, #4b5563)" }}>
                    <span style={{ color: "white" }}><IconGallery /></span>
                  </div>
                  <div className="text-center">
                    <p className="font-bold text-gray-900 text-base">Galeri</p>
                    <p className="text-xs text-gray-400 mt-0.5">Fotoğraf seç</p>
                  </div>
                </div>
              </div>
              <input type="file" accept="image/*" className="hidden" onChange={gorselSec} />
            </label>
          </div>

        </div>
      )}

      {/* ── İŞLENİYOR — Animasyonlu WINDOFORM Logo ─────────────────────── */}
      {adim === "processing" && (
        <div className="bg-white rounded-3xl shadow-md overflow-hidden">
          {/* Mavi üst bant */}
          <div className="h-1.5 w-full" style={{ background: wfBlue }} />

          <div className="p-10 flex flex-col items-center gap-8">
            {/* Taranan görsel */}
            {onizlemeGorsel && (
              // eslint-disable-next-line @next/next/no-img-element
              <img src={onizlemeGorsel} alt=""
                className="max-h-48 w-full object-contain rounded-2xl shadow-sm" />
            )}

            {/* Animasyonlu WINDOFORM logo */}
            <div className="relative flex flex-col items-center gap-5">
              <div className="relative inline-block">
                {/* Logo */}
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img
                  src="/windoform-logo.png"
                  alt="WINDOFORM"
                  className="h-12 w-auto object-contain logo-glow logo-pulse"
                  style={{ filter: "brightness(1)" }}
                />
                {/* Scan çizgisi */}
                <div
                  className="scan-bar absolute left-0 right-0 h-0.5 pointer-events-none"
                  style={{
                    background: `linear-gradient(90deg, transparent, ${wfBlue}, transparent)`,
                    top: 0,
                  }}
                />
              </div>

              <div className="text-center">
                <p className="font-bold text-gray-800 text-lg">Analiz ediliyor...</p>
                <p className="text-gray-400 text-sm mt-1">Yapay zeka el yazısını okuyor</p>
              </div>

              {/* Yükleme noktaları */}
              <div className="flex items-center gap-2">
                {[0,1,2].map(i => (
                  <div key={i} className={`w-2.5 h-2.5 rounded-full ${["dot-bounce-1","dot-bounce-2","dot-bounce-3"][i]}`}
                    style={{ background: wfBlue }} />
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── ONAY ────────────────────────────────────────────────────────── */}
      {adim === "confirm" && mevcutUrun && (
        <div className="space-y-3">
          {/* İlerleme */}
          <div className="bg-white rounded-2xl shadow-sm px-4 py-3">
            <div className="flex justify-between items-center mb-2">
              <span className="text-sm font-bold text-gray-700">{mevcutIndex + 1} / {toplamUrun}</span>
              {ralRenk && (
                <span className="text-xs font-mono font-bold text-white px-2.5 py-0.5 rounded-lg"
                  style={{ background: wfBlue }}>
                  RAL {ralRenk}
                </span>
              )}
            </div>
            <div className="w-full bg-gray-100 rounded-full h-1.5">
              <div className="h-1.5 rounded-full transition-all duration-500"
                style={{ width: `${(mevcutIndex / toplamUrun) * 100}%`, background: wfBlue }} />
            </div>
          </div>

          {/* Fişten okunan */}
          <div className="rounded-3xl p-5 space-y-2 border-2 border-amber-200 bg-amber-50">
            <p className="text-[11px] font-bold uppercase tracking-widest text-amber-500">📄 Fişten Okunan</p>
            <p className="text-xl font-bold text-gray-900 leading-tight">{mevcutUrun.urun_adi}</p>
            <p className="text-5xl font-black" style={{ color: wfBlue }}>{mevcutUrun.miktar}</p>
          </div>

          {/* Stok eşleşmesi */}
          {!duzenlemeAcik ? (
            <div className="bg-white rounded-3xl shadow-sm p-5 space-y-4">
              <p className="text-[11px] font-bold uppercase tracking-widest text-gray-400">🔍 Stok Eşleşmesi</p>
              {duzenStok ? (
                <div className="rounded-2xl p-4 border-2" style={{ borderColor: wfBlue, background: "#EEF2F8" }}>
                  <p className="font-mono font-bold text-sm" style={{ color: wfBlue }}>{duzenStok.stok_kodu}</p>
                  <p className="font-bold text-gray-900 text-base mt-1 leading-snug">{duzenStok.stok_adi}</p>
                </div>
              ) : (
                <div className="bg-red-50 border-2 border-red-200 rounded-2xl p-4 text-center">
                  <p className="text-red-600 font-bold text-sm">Eşleşen ürün bulunamadı</p>
                </div>
              )}
              {stokOneriler.length > 1 && (
                <div className="space-y-1.5">
                  <p className="text-xs text-gray-400 font-medium">Diğer öneriler:</p>
                  {stokOneriler.slice(1,4).map(s => (
                    <button key={s.stok_kodu} onClick={() => setDuzenStok(s)}
                      className={`w-full text-left text-sm rounded-xl px-4 py-2.5 border transition-colors ${
                        duzenStok?.stok_kodu === s.stok_kodu ? "border-blue-300 bg-blue-50 font-bold" : "bg-gray-50 border-gray-200"}`}>
                      <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span>
                      <span className="ml-2 text-gray-700">{s.stok_adi}</span>
                    </button>
                  ))}
                </div>
              )}
              <div className="flex items-center justify-between pt-3 border-t border-gray-100">
                <div>
                  <p className="text-xs text-gray-400 font-medium">Miktar</p>
                  <p className="text-4xl font-black" style={{ color: wfBlue }}>{duzenMiktar}</p>
                </div>
                <button onClick={() => setDuzenlemeAcik(true)}
                  className="bg-gray-100 active:bg-gray-200 text-gray-700 font-bold px-5 py-3 rounded-2xl text-sm">
                  ✏️ Düzenle
                </button>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-3xl shadow-sm p-5 space-y-4">
              <p className="text-[11px] font-bold uppercase tracking-widest text-gray-400">✏️ Düzenleme</p>
              <div>
                <label className="text-xs font-bold text-gray-400 block mb-2 uppercase tracking-wider">Miktar</label>
                <input type="number" value={duzenMiktar} onChange={e => setDuzenMiktar(e.target.value)} autoFocus
                  className="w-full border-2 rounded-2xl p-4 text-3xl font-bold text-center focus:outline-none"
                  style={{ borderColor: wfBlue }} />
              </div>
              <div>
                <label className="text-xs font-bold text-gray-400 block mb-2 uppercase tracking-wider">Ürün Ara</label>
                <input type="text" value={stokArama} onChange={e => stokAramaGuncelle(e.target.value)}
                  placeholder="Ürün adı veya stok kodu..."
                  className="w-full border-2 border-gray-200 rounded-2xl p-3.5 text-sm focus:outline-none focus:border-blue-300" />
                {stokOneriler.length > 0 && (
                  <div className="mt-2 space-y-1 max-h-52 overflow-y-auto">
                    {stokOneriler.map(s => (
                      <button key={s.stok_kodu}
                        onClick={() => { setDuzenStok(s); setStokArama(s.stok_adi); setStokOneriler([]); }}
                        className={`w-full text-left text-sm rounded-xl px-4 py-3 border transition-colors ${
                          duzenStok?.stok_kodu === s.stok_kodu ? "bg-blue-50 border-blue-300 font-bold" : "bg-gray-50 border-gray-200"}`}>
                        <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span>
                        <span className="ml-2">{s.stok_adi}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <button onClick={() => setDuzenlemeAcik(false)}
                className="w-full text-white font-bold py-4 rounded-2xl text-base active:opacity-90"
                style={{ background: wfBlue }}>
                💾 Kaydet
              </button>
            </div>
          )}

          {!duzenlemeAcik && (
            <div className="grid grid-cols-2 gap-3">
              {/* ATLA — weißer Premium-Button */}
              <button onClick={atla}
                className="group relative overflow-hidden bg-white rounded-3xl shadow-sm border border-gray-200 active:scale-95 transition-transform">
                <div className="absolute inset-0 bg-gray-50 opacity-0 group-active:opacity-100 transition-opacity" />
                <div className="relative flex flex-col items-center justify-center gap-2 py-7 px-4">
                  <div className="w-11 h-11 rounded-2xl bg-gray-100 flex items-center justify-center">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6b7280" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <polygon points="5 4 15 12 5 20 5 4"/><line x1="19" y1="5" x2="19" y2="19"/>
                    </svg>
                  </div>
                  <span className="font-bold text-gray-600 text-base">Atla</span>
                </div>
              </button>

              {/* EVET — grüner Premium-Button */}
              <button onClick={onayla}
                className="group relative overflow-hidden rounded-3xl shadow-md active:scale-95 transition-transform"
                style={{ background: "linear-gradient(145deg, #16a34a, #15803d)" }}>
                <div className="absolute inset-0 bg-white opacity-0 group-active:opacity-10 transition-opacity" />
                <div className="relative flex flex-col items-center justify-center gap-2 py-7 px-4">
                  <div className="w-11 h-11 rounded-2xl bg-white/20 flex items-center justify-center">
                    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="20 6 9 17 4 12"/>
                    </svg>
                  </div>
                  <span className="font-bold text-white text-base">Onayla</span>
                </div>
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── BAŞKA SAYFA ─────────────────────────────────────────────────── */}
      {adim === "more_pages" && (
        <div className="bg-white rounded-3xl shadow-md overflow-hidden">
          <div className="h-1.5" style={{ background: wfBlue }} />
          <div className="px-6 py-8 flex flex-col items-center text-center gap-6">
            {/* Logo in weißer Pille */}
            <div className="bg-gray-50 border border-gray-100 rounded-2xl px-5 py-3">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src="/windoform-logo.png" alt="WINDOFORM" className="h-7 object-contain" />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-gray-900">Başka Sayfa Var Mı?</h2>
              <p className="text-gray-400 mt-1.5 text-sm">
                Şu ana kadar <span className="font-bold" style={{ color: wfBlue }}>{tumOnaylananlar.filter(i => !i.skipped).length}</span> ürün onaylandı
              </p>
            </div>
            <div className="grid grid-cols-2 gap-3 w-full">
              {/* Evet */}
              <button
                onClick={() => { setTaramaKonusu(null); setSayfaOnaylananlar([]); setMevcutIndex(0); setOnizlemeGorsel(null); setRalRenk(""); setAdimGecmisi([]); gidim("ral"); }}
                className="group relative overflow-hidden rounded-3xl shadow-md active:scale-95 transition-transform"
                style={{ background: "linear-gradient(145deg, #16a34a, #15803d)" }}>
                <div className="flex flex-col items-center justify-center gap-2.5 py-7 px-4">
                  <div className="w-12 h-12 rounded-2xl bg-white/20 flex items-center justify-center">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="20 6 9 17 4 12"/>
                    </svg>
                  </div>
                  <span className="font-bold text-white text-base">Evet, Devam</span>
                </div>
              </button>
              {/* Hayır */}
              <button
                onClick={() => gidim("done")}
                className="group relative overflow-hidden rounded-3xl shadow-md active:scale-95 transition-transform bg-white border-2 border-gray-200">
                <div className="flex flex-col items-center justify-center gap-2.5 py-7 px-4">
                  <div className="w-12 h-12 rounded-2xl bg-gray-100 flex items-center justify-center">
                    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#6b7280" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/>
                    </svg>
                  </div>
                  <span className="font-bold text-gray-700 text-base">Bitir</span>
                </div>
              </button>
            </div>
            <p className="text-gray-300 text-xs">"Bitir" seçince Excel hazırlanır</p>
          </div>
        </div>
      )}

      {/* ── TAMAMLANDI ──────────────────────────────────────────────────── */}
      {adim === "done" && (
        <div className="space-y-3">

          {/* Hero-Karte — Logo in weißer Pille, kein Filter */}
          <div className="rounded-3xl overflow-hidden shadow-lg"
            style={{ background: `linear-gradient(135deg, ${wfBlue} 0%, #1e3d6e 100%)` }}>
            <div className="px-6 pt-7 pb-8 flex flex-col items-center text-center gap-5">
              {/* Logo in weißer Pille */}
              <div className="bg-white rounded-xl px-4 py-2 shadow-sm">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img src="/windoform-logo.png" alt="WINDOFORM" className="h-6 object-contain" />
              </div>
              <div>
                <p className="text-white/60 text-xs font-semibold uppercase tracking-widest">Tamamlandı</p>
                <p className="text-white font-black mt-1" style={{ fontSize: "clamp(2.5rem,10vw,3.5rem)", lineHeight: 1 }}>
                  {onaylananSayisi}
                </p>
                <p className="text-white/80 text-lg font-medium mt-0.5">Ürün Onaylandı</p>
                <p className="text-white/50 text-sm mt-2">{tarih}</p>
              </div>
            </div>
          </div>

          {/* Ürün listesi */}
          <div className="bg-white rounded-3xl shadow-sm overflow-hidden">
            <div className="px-5 pt-4 pb-2">
              <p className="text-[11px] font-bold uppercase tracking-widest text-gray-400">Onaylanan Ürünler</p>
            </div>
            <div className="divide-y divide-gray-50 max-h-72 overflow-y-auto">
              {tumOnaylananlar.map((u, idx) => (
                <div key={idx}
                  className={`flex items-center gap-3 px-5 py-3.5 ${u.skipped ? "opacity-40" : ""}`}>
                  {/* İkon */}
                  <div className={`w-8 h-8 rounded-xl flex items-center justify-center flex-shrink-0 ${
                    u.skipped ? "bg-gray-100" : "bg-green-50"}`}>
                    {u.skipped ? (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <polygon points="5 4 15 12 5 20 5 4"/><line x1="19" y1="5" x2="19" y2="19"/>
                      </svg>
                    ) : (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                        <polyline points="20 6 9 17 4 12"/>
                      </svg>
                    )}
                  </div>
                  {/* Metin */}
                  <div className="flex-1 min-w-0">
                    <p className="font-semibold text-sm text-gray-900 truncate leading-tight">
                      {u.confirmed_stok?.stok_adi || u.original_urun_adi}
                    </p>
                    {u.confirmed_stok && (
                      <p className="text-[11px] font-mono text-gray-400 mt-0.5">{u.confirmed_stok.stok_kodu}</p>
                    )}
                  </div>
                  {/* Miktar */}
                  {!u.skipped && (
                    <span className="font-black text-lg flex-shrink-0" style={{ color: wfBlue }}>
                      {u.confirmed_miktar}
                    </span>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* Excel İndir — Premium blauer Button */}
          <button onClick={excelIndir}
            className="w-full relative overflow-hidden rounded-3xl shadow-lg active:scale-95 transition-transform"
            style={{ background: `linear-gradient(145deg, ${wfBlue}, #1e3d6e)` }}>
            <div className="flex items-center justify-center gap-3 py-5 px-6">
              <div className="w-10 h-10 rounded-2xl bg-white/20 flex items-center justify-center">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
                  <polyline points="7 10 12 15 17 10"/>
                  <line x1="12" y1="15" x2="12" y2="3"/>
                </svg>
              </div>
              <div className="text-left">
                <p className="text-white font-bold text-lg leading-tight">Excel İndir</p>
                <p className="text-white/60 text-xs">Boş Üsk formatında</p>
              </div>
            </div>
          </button>

          {/* OneDrive — Button */}
          {odDurum !== "success" ? (
            <button onClick={oneDriveYukle} disabled={odDurum === "loading"}
              className="w-full relative overflow-hidden rounded-3xl shadow-sm border-2 active:scale-95 transition-transform disabled:opacity-60"
              style={{ borderColor: "#0078d4", background: "#fff" }}>
              <div className="flex items-center justify-center gap-3 py-5 px-6">
                <div className="w-10 h-10 rounded-2xl flex items-center justify-center"
                  style={{ background: "#e8f1fb" }}>
                  {odDurum === "loading" ? (
                    <div className="w-5 h-5 rounded-full border-2 border-t-transparent animate-spin"
                      style={{ borderColor: "#0078d4", borderTopColor: "transparent" }} />
                  ) : (
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0078d4" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1"/>
                      <polyline points="16 6 12 2 8 6"/>
                      <line x1="12" y1="2" x2="12" y2="15"/>
                    </svg>
                  )}
                </div>
                <div className="text-left">
                  <p className="font-bold text-base leading-tight" style={{ color: "#0078d4" }}>
                    {odDurum === "loading" ? "Yükleniyor..." : "OneDrive'a Kaydet"}
                  </p>
                  <p className="text-xs mt-0.5" style={{ color: "#666" }}>
                    Excel + {taramaFotolari.length} fotoğraf → WINDOFORM/Uretim
                  </p>
                </div>
              </div>
            </button>
          ) : (
            <div className="w-full rounded-3xl py-4 px-6 text-center text-sm font-medium text-green-700 bg-green-50 border border-green-200">
              {odMesaj}
            </div>
          )}
          {odDurum === "error" && odMesaj && (
            <div className="text-xs text-red-600 text-center px-2">{odMesaj}</div>
          )}

          {/* Yeni Fiş — weißer Button */}
          <button onClick={yenidenBaslat}
            className="w-full bg-white border border-gray-200 rounded-3xl shadow-sm active:bg-gray-50 active:scale-95 transition-transform">
            <div className="flex items-center justify-center gap-3 py-4 px-6">
              <div className="w-9 h-9 rounded-xl bg-gray-100 flex items-center justify-center">
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="#6b7280" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/>
                  <path d="M3 3v5h5"/>
                </svg>
              </div>
              <span className="font-bold text-gray-600 text-base">Yeni Fiş Tara</span>
            </div>
          </button>
        </div>
      )}

      {/* ── GÜNCELLE ────────────────────────────────────────────────────── */}
      {adim === "update" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-sm overflow-hidden">
            <div className="h-1.5" style={{ background: wfBlue }} />
            <div className="p-6 space-y-4">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src="/windoform-logo.png" alt="WINDOFORM" className="h-7 object-contain" />
              <div>
                <h2 className="text-lg font-bold text-gray-900">Stok Kartını Güncelle</h2>
                <p className="text-sm text-gray-400 mt-0.5">{stokData.length.toLocaleString("tr-TR")} ürün kayıtlı</p>
              </div>
              <p className="text-sm text-gray-500">Yalnızca <strong>yeni stok kodları</strong> eklenir.</p>
              <label className="block cursor-pointer">
                <div className="border-2 border-dashed rounded-2xl p-8 text-center active:bg-blue-50 transition-all"
                  style={{ borderColor: wfBlue }}>
                  <div className="flex flex-col items-center gap-2">
                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke={wfBlue} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/>
                      <line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
                    </svg>
                    <p className="font-bold text-base" style={{ color: wfBlue }}>Excel Dosyası Seç</p>
                    <p className="text-gray-400 text-xs">Stok Kart Kayıtları.xlsx</p>
                  </div>
                </div>
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={stokGuncelle} />
              </label>
              {guncellemeMetni && (
                <div className={`rounded-2xl p-4 text-sm font-medium ${
                  guncellemeMetni.startsWith("✅") ? "bg-green-50 text-green-800 border border-green-200" :
                  guncellemeMetni.startsWith("❌") ? "bg-red-50 text-red-800 border border-red-200" : "bg-gray-50 text-gray-600"}`}>
                  {guncellemeMetni}
                </div>
              )}

              {/* Öğrenilenleri sıfırla */}
              <div className="pt-2 border-t border-gray-100">
                <p className="text-xs text-gray-400 mb-2">
                  {(() => {
                    try { return Object.keys(JSON.parse(localStorage.getItem(OGRENIM_KEY) || "{}")).length; } catch { return 0; }
                  })()} öğrenilmiş düzeltme kayıtlı
                </p>
                <button
                  onClick={() => { localStorage.removeItem(OGRENIM_KEY); setGuncellemeMetni("✅ Öğrenilmiş düzeltmeler sıfırlandı."); }}
                  className="text-xs text-red-400 underline">
                  Öğrenilenleri Sıfırla
                </button>
              </div>
            </div>
          </div>
          <button onClick={() => setAdim("scan")}
            className="w-full bg-white border border-gray-200 text-gray-500 font-bold py-4 rounded-3xl text-base shadow-sm active:bg-gray-50 flex items-center justify-center gap-2">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><polyline points="15 18 9 12 15 6"/></svg>
            Geri Dön
          </button>
        </div>
      )}
    </div>
  );
}
