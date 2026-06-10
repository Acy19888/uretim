"use client";

import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { findMatches } from "@/lib/fuzzy";
import type { StokItem, ScannedItem, ConfirmedItem, ScanResult } from "@/lib/types";
import bundledStok from "@/lib/stokData.json";

type Step = "scan" | "processing" | "confirm" | "more_pages" | "done" | "update";

// ── WINDOFORM W-mark SVG ───────────────────────────────────────────────────
function WMark({
  size = 56,
  color = "#2B5597",
  animated = false,
}: {
  size?: number;
  color?: string;
  animated?: boolean;
}) {
  const w = Math.round(size * 1.17);
  return (
    <svg width={w} height={size} viewBox="0 0 56 48" fill="none" xmlns="http://www.w3.org/2000/svg"
      className={animated ? "wf-logo-pulse" : ""}>
      {/* Left arm */}
      <path
        className={animated ? "wf-logo-left" : ""}
        d="M0,0 L7,0 L16,48 L7,48 Z"
        fill={color}
      />
      {/* Left-center arm */}
      <path
        className={animated ? "wf-logo-left" : ""}
        d="M7,0 L20,0 L16,48 L3,48 Z"
        fill={color} opacity="0.9"
      />
      {/* Center peak */}
      <path
        className={animated ? "wf-logo-center" : ""}
        d="M0,0 L56,0 L47,48 L28,14 L9,48 L0,0 Z
           M7,0 L28,36 L49,0 Z"
        fill={color}
        fillRule="evenodd"
      />
      {/* Right-center arm */}
      <path
        className={animated ? "wf-logo-right" : ""}
        d="M36,0 L49,0 L53,48 L40,48 Z"
        fill={color} opacity="0.9"
      />
      {/* Right arm */}
      <path
        className={animated ? "wf-logo-right" : ""}
        d="M49,0 L56,0 L49,48 L40,48 Z"
        fill={color}
      />
    </svg>
  );
}

// Clean single-path W-mark
function WMarkClean({
  size = 56,
  color = "#2B5597",
  animated = false,
  className = "",
}: {
  size?: number;
  color?: string;
  animated?: boolean;
  className?: string;
}) {
  return (
    <svg
      width={Math.round(size * 1.17)}
      height={size}
      viewBox="0 0 56 48"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      className={`${animated ? "wf-logo-pulse" : ""} ${className}`}
    >
      <path d="M0,0 L9,48 L28,12 L47,48 L56,0 L49,0 L28,36 L7,0 Z" fill={color} />
      {/* inner W cutout to create the hollow W effect */}
      <path d="M9,48 L14,48 L28,24 L42,48 L47,48 L28,12 Z" fill="var(--wf-bg, #EEF2F8)" />
    </svg>
  );
}

function toBase64(file: File): Promise<string> {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.readAsDataURL(file);
    r.onload  = () => res((r.result as string).split(",")[1]);
    r.onerror = rej;
  });
}

function bugunStr(): string {
  const d = new Date();
  return `${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}.${d.getFullYear()}`;
}

function stokYukle(): StokItem[] {
  const base = bundledStok as StokItem[];
  try {
    const raw = localStorage.getItem("stok_guncellemeler");
    if (!raw) return base;
    const guncellemeler: StokItem[] = JSON.parse(raw);
    const map = new Map(base.map(i => [i.stok_kodu, i]));
    for (const g of guncellemeler) map.set(g.stok_kodu, g);
    return Array.from(map.values());
  } catch { return base; }
}

// ── Component ─────────────────────────────────────────────────────────────
export default function Home() {
  const [adim, setAdim]                       = useState<Step>("scan");
  const [stokData, setStokData]               = useState<StokItem[]>([]);
  const [tumOnaylananlar, setTumOnaylananlar]  = useState<ConfirmedItem[]>([]);
  const [isEmriNo]                            = useState("");
  const [tarih]                               = useState(bugunStr());
  const [taramaKonusu, setTaramaKonusu]       = useState<ScanResult | null>(null);
  const [mevcutIndex, setMevcutIndex]         = useState(0);
  const [sayfaOnaylananlar, setSayfaOnaylananlar] = useState<ConfirmedItem[]>([]);
  const [onizlemeGorsel, setOnizlemeGorsel]   = useState<string | null>(null);
  const [hata, setHata]                       = useState<string | null>(null);
  const [ralRenk, setRalRenk]                 = useState("");
  const [duzenMiktar, setDuzenMiktar]         = useState("");
  const [duzenStok, setDuzenStok]             = useState<StokItem | null>(null);
  const [stokArama, setStokArama]             = useState("");
  const [stokOneriler, setStokOneriler]       = useState<StokItem[]>([]);
  const [duzenlemeAcik, setDuzenlemeAcik]     = useState(false);
  const [guncellemeMetni, setGuncellemeMetni] = useState<string | null>(null);

  useEffect(() => { setStokData(stokYukle()); }, []);

  // ── Tarama ──────────────────────────────────────────────────────────────
  async function gorselSec(e: React.ChangeEvent<HTMLInputElement>) {
    const dosya = e.target.files?.[0];
    if (!dosya) return;
    setHata(null);
    setOnizlemeGorsel(URL.createObjectURL(dosya));
    setAdim("processing");
    try {
      const b64 = await toBase64(dosya);
      const res = await fetch("/api/scan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ imageBase64: b64, mimeType: dosya.type }),
      });
      if (!res.ok) throw new Error((await res.json()).error || "Tarama başarısız");
      const sonuc: ScanResult = await res.json();
      if (!sonuc.items?.length) throw new Error("Hiçbir ürün algılanamadı. Tekrar deneyin.");
      if (sonuc.ral_renk) setRalRenk(sonuc.ral_renk);
      setTaramaKonusu(sonuc);
      setMevcutIndex(0);
      setSayfaOnaylananlar([]);
      setAdim("confirm");
    } catch (err) { setHata(String(err)); setAdim("scan"); }
  }

  // ── Onay ────────────────────────────────────────────────────────────────
  const satirHazirla = useCallback((urun: ScannedItem) => {
    setDuzenMiktar(urun.miktar || "");
    setStokArama(urun.urun_adi || "");
    setDuzenlemeAcik(false);
    const aramaMetni = ralRenk ? `${ralRenk} ${urun.urun_adi}` : urun.urun_adi;
    const eslesimler = findMatches(aramaMetni, stokData, 5);
    setStokOneriler(eslesimler);
    setDuzenStok(eslesimler[0] || null);
  }, [stokData, ralRenk]);

  useEffect(() => {
    if (adim === "confirm" && taramaKonusu && mevcutIndex < taramaKonusu.items.length)
      satirHazirla(taramaKonusu.items[mevcutIndex]);
  }, [mevcutIndex, adim, taramaKonusu, satirHazirla]);

  function onayla() {
    const urun = taramaKonusu!.items[mevcutIndex];
    ilerle([...sayfaOnaylananlar, {
      original_urun_adi: urun.urun_adi, original_miktar: urun.miktar,
      confirmed_stok: duzenStok, confirmed_miktar: parseFloat(duzenMiktar.replace(",",".")) || 0,
      skipped: false,
    }]);
  }

  function atla() {
    const urun = taramaKonusu!.items[mevcutIndex];
    ilerle([...sayfaOnaylananlar, {
      original_urun_adi: urun.urun_adi, original_miktar: urun.miktar,
      confirmed_stok: null, confirmed_miktar: 0, skipped: true,
    }]);
  }

  function ilerle(guncellenmis: ConfirmedItem[]) {
    setSayfaOnaylananlar(guncellenmis);
    if (mevcutIndex + 1 < taramaKonusu!.items.length) {
      setMevcutIndex(mevcutIndex + 1);
    } else {
      setTumOnaylananlar(prev => [...prev, ...guncellenmis]);
      setAdim("more_pages");
    }
  }

  function stokAramaGuncelle(q: string) {
    setStokArama(q);
    setStokOneriler(q.length >= 2 ? findMatches(q, stokData, 8) : []);
  }

  // ── Stok güncelle ───────────────────────────────────────────────────────
  async function stokGuncelle(e: React.ChangeEvent<HTMLInputElement>) {
    const dosya = e.target.files?.[0];
    if (!dosya) return;
    setGuncellemeMetni("Okunuyor...");
    try {
      const buf   = await dosya.arrayBuffer();
      const wb    = XLSX.read(buf, { type: "array" });
      const ws    = wb.Sheets[wb.SheetNames[0]];
      const satirlar = XLSX.utils.sheet_to_json<Record<string,string>>(ws, { raw: false, defval: "" });
      const yeniUrunler: StokItem[] = satirlar.flatMap(satir => {
        const kodu  = (satir["STOK_KODU"] || satir["Stok Kodu"] || "").trim();
        const adi   = (satir["STOK_ADI"]  || satir["Stok Adı"]  || "").trim();
        const cesit = (satir["Çeşit"] || "").trim();
        if (!kodu || !adi) return [];
        return [{ stok_kodu: kodu, stok_adi: adi, cesit }];
      });
      const mevcutKodlar = new Set((bundledStok as StokItem[]).map(i => i.stok_kodu));
      const gercektenYeni = yeniUrunler.filter(i => !mevcutKodlar.has(i.stok_kodu));
      const eskiRaw = localStorage.getItem("stok_guncellemeler");
      const eski: StokItem[] = eskiRaw ? JSON.parse(eskiRaw) : [];
      const map = new Map(eski.map(i => [i.stok_kodu, i]));
      for (const y of gercektenYeni) map.set(y.stok_kodu, y);
      localStorage.setItem("stok_guncellemeler", JSON.stringify(Array.from(map.values())));
      setStokData(stokYukle());
      setGuncellemeMetni(`✅ ${gercektenYeni.length} yeni ürün eklendi. Toplam: ${yeniUrunler.length}`);
    } catch (err) { setGuncellemeMetni("❌ Hata: " + String(err)); }
  }

  // ── Excel export ────────────────────────────────────────────────────────
  async function excelIndir() {
    setHata(null);
    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
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
    setMevcutIndex(0); setOnizlemeGorsel(null); setRalRenk(""); setAdim("scan");
  }

  const mevcutUrun      = taramaKonusu?.items[mevcutIndex];
  const toplamUrun      = taramaKonusu?.items.length || 0;
  const onaylananSayisi = tumOnaylananlar.filter(i => !i.skipped).length;

  // ── Styles ──────────────────────────────────────────────────────────────
  const wfBlue     = "#2B5597";
  const wfBlueDark = "#1e3d6e";

  return (
    <div className="space-y-4 pb-12">

      {/* Hata */}
      {hata && (
        <div className="bg-red-50 border border-red-200 text-red-800 rounded-2xl p-4 flex gap-3 items-start shadow-sm">
          <span className="text-xl mt-0.5">⚠️</span>
          <div className="flex-1 text-sm leading-relaxed">{hata}</div>
          <button onClick={() => setHata(null)} className="text-red-300 text-xl font-bold leading-none">✕</button>
        </div>
      )}

      {/* ── TARAMA ────────────────────────────────────────────────────── */}
      {adim === "scan" && (
        <div className="space-y-3 wf-fade-up">
          {/* Üst bilgi kartı */}
          <div className="bg-white rounded-3xl shadow-sm border border-gray-100 p-4 flex justify-between items-center">
            <div>
              <p className="text-xs font-semibold uppercase tracking-wider" style={{ color: wfBlue }}>
                Yarı Mamul Çıkış
              </p>
              {tumOnaylananlar.length > 0 && (
                <p className="text-sm text-green-600 font-medium mt-0.5">
                  ✓ {onaylananSayisi} ürün onaylandı
                </p>
              )}
            </div>
            <div className="text-right">
              <p className="text-xs text-gray-400">Stok</p>
              <p className="text-lg font-bold" style={{ color: wfBlue }}>
                {stokData.length.toLocaleString("tr-TR")}
              </p>
            </div>
          </div>

          {/* Kamera butonu */}
          <label className="block cursor-pointer">
            <div
              className="rounded-3xl p-10 text-center shadow-md active:opacity-90 transition-all"
              style={{ background: `linear-gradient(135deg, ${wfBlue} 0%, ${wfBlueDark} 100%)` }}
            >
              <div className="text-6xl mb-3">📷</div>
              <p className="text-white text-2xl font-bold tracking-wide">Kamera</p>
              <p className="text-white/60 text-sm mt-1">Fotoğraf çek</p>
            </div>
            <input type="file" accept="image/*" capture="environment" className="hidden" onChange={gorselSec} />
          </label>

          {/* Galeriden seç */}
          <label className="block cursor-pointer">
            <div className="bg-white border border-gray-200 rounded-3xl p-7 text-center active:bg-gray-50 transition-all shadow-sm">
              <div className="text-4xl mb-2">🖼️</div>
              <p className="text-base font-bold text-gray-700">Galeriden Seç</p>
            </div>
            <input type="file" accept="image/*" className="hidden" onChange={gorselSec} />
          </label>

          {/* Güncelle */}
          <button
            onClick={() => { setAdim("update"); setGuncellemeMetni(null); }}
            className="w-full bg-white border border-gray-200 text-gray-400 font-medium py-4 rounded-2xl text-sm shadow-sm active:bg-gray-50"
          >
            🔄 Stok Kartını Güncelle
          </button>
        </div>
      )}

      {/* ── İŞLENİYOR ──────────────────────────────────────────────────── */}
      {adim === "processing" && (
        <div className="bg-white rounded-3xl shadow-md p-10 text-center space-y-8">
          {onizlemeGorsel && (
            // eslint-disable-next-line @next/next/no-img-element
            <img src={onizlemeGorsel} alt="" className="max-h-52 mx-auto rounded-2xl shadow object-contain" />
          )}

          {/* Animated WINDOFORM W-mark */}
          <div className="flex flex-col items-center gap-4">
            <WMarkClean size={72} color={wfBlue} animated={true} />
            <div>
              <p className="text-xl font-bold" style={{ color: wfBlue }}>Analiz ediliyor...</p>
              <p className="text-gray-400 text-sm mt-1">Yapay zeka el yazısını okuyor</p>
            </div>
          </div>

          {/* Loading dots */}
          <div className="flex justify-center gap-2">
            {[0,1,2].map(i => (
              <div
                key={i}
                className="w-2 h-2 rounded-full"
                style={{
                  background: wfBlue,
                  animation: `wPulse 1.2s ease-in-out ${i * 0.2}s infinite`,
                  opacity: 0.6,
                }}
              />
            ))}
          </div>
        </div>
      )}

      {/* ── ONAY ────────────────────────────────────────────────────────── */}
      {adim === "confirm" && mevcutUrun && (
        <div className="space-y-3">
          {/* İlerleme */}
          <div className="bg-white rounded-2xl shadow-sm p-4">
            <div className="flex justify-between items-center mb-2">
              <span className="text-sm font-bold text-gray-700">{mevcutIndex + 1} / {toplamUrun}</span>
              <div className="flex items-center gap-2">
                {ralRenk && (
                  <span className="text-xs font-mono font-bold text-white px-2 py-0.5 rounded-lg"
                    style={{ background: wfBlue }}>
                    RAL {ralRenk}
                  </span>
                )}
              </div>
            </div>
            <div className="w-full bg-gray-100 rounded-full h-2">
              <div
                className="h-2 rounded-full transition-all duration-500"
                style={{ width: `${((mevcutIndex) / toplamUrun) * 100}%`, background: wfBlue }}
              />
            </div>
          </div>

          {/* Fişten okunan */}
          <div className="rounded-3xl p-5 space-y-2 border-2 border-amber-200 bg-amber-50">
            <p className="text-[11px] font-bold uppercase tracking-widest text-amber-600">📄 Fişten Okunan</p>
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
                  <p className="text-red-600 font-bold">Eşleşen ürün bulunamadı</p>
                </div>
              )}

              {stokOneriler.length > 1 && (
                <div className="space-y-1.5">
                  <p className="text-xs text-gray-400 font-medium">Diğer öneriler:</p>
                  {stokOneriler.slice(1, 4).map(s => (
                    <button key={s.stok_kodu} onClick={() => setDuzenStok(s)}
                      className={`w-full text-left text-sm rounded-xl px-4 py-2.5 border transition-all ${
                        duzenStok?.stok_kodu === s.stok_kodu
                          ? "border-blue-300 bg-blue-50 font-bold"
                          : "bg-gray-50 border-gray-200 active:bg-blue-50"
                      }`}>
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
                <label className="text-xs font-bold text-gray-500 block mb-2 uppercase tracking-wider">Miktar</label>
                <input type="number" value={duzenMiktar} onChange={e => setDuzenMiktar(e.target.value)} autoFocus
                  className="w-full border-2 rounded-2xl p-4 text-3xl font-bold text-center focus:outline-none"
                  style={{ borderColor: wfBlue }} />
              </div>
              <div>
                <label className="text-xs font-bold text-gray-500 block mb-2 uppercase tracking-wider">Ürün Ara</label>
                <input type="text" value={stokArama} onChange={e => stokAramaGuncelle(e.target.value)}
                  placeholder="Ürün adı veya stok kodu..."
                  className="w-full border-2 border-gray-200 rounded-2xl p-3.5 text-sm focus:outline-none focus:border-blue-400" />
                {stokOneriler.length > 0 && (
                  <div className="mt-2 space-y-1 max-h-52 overflow-y-auto">
                    {stokOneriler.map(s => (
                      <button key={s.stok_kodu}
                        onClick={() => { setDuzenStok(s); setStokArama(s.stok_adi); setStokOneriler([]); }}
                        className={`w-full text-left text-sm rounded-xl px-4 py-3 border transition-colors ${
                          duzenStok?.stok_kodu === s.stok_kodu
                            ? "bg-blue-50 border-blue-300 font-bold"
                            : "bg-gray-50 border-gray-200 active:bg-blue-50"
                        }`}>
                        <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span>
                        <span className="ml-2">{s.stok_adi}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <button onClick={() => setDuzenlemeAcik(false)}
                className="w-full text-white font-bold py-4 rounded-2xl text-base"
                style={{ background: wfBlue }}>
                💾 Kaydet
              </button>
            </div>
          )}

          {!duzenlemeAcik && (
            <div className="grid grid-cols-2 gap-3">
              <button onClick={atla}
                className="bg-gray-200 active:bg-gray-300 text-gray-700 font-bold py-6 rounded-3xl text-xl">
                ⏭ Atla
              </button>
              <button onClick={onayla}
                className="text-white font-bold py-6 rounded-3xl text-2xl shadow-lg active:opacity-90"
                style={{ background: "#16a34a" }}>
                ✅ Evet
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── BAŞKA SAYFA ─────────────────────────────────────────────────── */}
      {adim === "more_pages" && (
        <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-6">
          <WMarkClean size={52} color={wfBlue} className="mx-auto" />
          <div>
            <h2 className="text-2xl font-bold text-gray-900">Başka Sayfa Var Mı?</h2>
            <p className="text-gray-400 mt-1">
              {tumOnaylananlar.filter(i => !i.skipped).length} ürün onaylandı
            </p>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <button
              onClick={() => { setTaramaKonusu(null); setSayfaOnaylananlar([]); setMevcutIndex(0); setOnizlemeGorsel(null); setAdim("scan"); }}
              className="text-white font-bold py-7 rounded-3xl text-2xl shadow-lg"
              style={{ background: "#16a34a" }}>
              ✅ Evet
            </button>
            <button onClick={() => setAdim("done")}
              className="text-white font-bold py-7 rounded-3xl text-2xl shadow-lg"
              style={{ background: "#dc2626" }}>
              ❌ Hayır
            </button>
          </div>
          <p className="text-gray-400 text-xs">Hayır → Excel oluşturulur</p>
        </div>
      )}

      {/* ── TAMAMLANDI ──────────────────────────────────────────────────── */}
      {adim === "done" && (
        <div className="space-y-4">
          <div className="rounded-3xl p-8 text-center space-y-4 text-white shadow-lg"
            style={{ background: `linear-gradient(135deg, ${wfBlue} 0%, ${wfBlueDark} 100%)` }}>
            <WMarkClean size={48} color="white" className="mx-auto" />
            <h2 className="text-3xl font-bold">Tamamlandı!</h2>
            <p className="opacity-80 text-lg">
              <strong>{onaylananSayisi}</strong> ürün — {tarih}
            </p>
          </div>

          <div className="bg-white rounded-3xl shadow-sm p-5 space-y-2">
            <h3 className="font-bold text-gray-700 text-sm uppercase tracking-wider">Onaylanan Ürünler</h3>
            <div className="space-y-1.5 max-h-64 overflow-y-auto">
              {tumOnaylananlar.map((urun, idx) => (
                <div key={idx}
                  className={`flex items-center gap-3 p-3 rounded-2xl ${urun.skipped ? "bg-gray-50 text-gray-400" : "bg-blue-50"}`}>
                  <span className="text-base">{urun.skipped ? "⏭" : "✅"}</span>
                  <div className="flex-1 min-w-0">
                    <p className="font-medium truncate text-sm">
                      {urun.confirmed_stok?.stok_adi || urun.original_urun_adi}
                    </p>
                    {urun.confirmed_stok &&
                      <p className="text-xs font-mono text-gray-400">{urun.confirmed_stok.stok_kodu}</p>}
                  </div>
                  {!urun.skipped &&
                    <span className="font-bold text-lg whitespace-nowrap" style={{ color: wfBlue }}>
                      {urun.confirmed_miktar}
                    </span>}
                </div>
              ))}
            </div>
          </div>

          <button onClick={excelIndir}
            className="w-full text-white font-bold py-6 rounded-3xl text-xl shadow-xl active:opacity-90"
            style={{ background: wfBlue }}>
            📥 Excel İndir
          </button>

          <button onClick={yenidenBaslat}
            className="w-full bg-white border border-gray-200 text-gray-600 font-bold py-4 rounded-3xl text-base shadow-sm">
            🔄 Yeni Fiş Tara
          </button>
        </div>
      )}

      {/* ── GÜNCELLE ────────────────────────────────────────────────────── */}
      {adim === "update" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-sm p-6 space-y-4">
            <div className="flex items-center gap-3">
              <WMarkClean size={32} color={wfBlue} />
              <div>
                <h2 className="text-lg font-bold" style={{ color: wfBlue }}>Stok Kartını Güncelle</h2>
                <p className="text-xs text-gray-400">{stokData.length.toLocaleString("tr-TR")} ürün kayıtlı</p>
              </div>
            </div>
            <p className="text-sm text-gray-500">
              Yalnızca <strong>yeni stok kodları</strong> eklenir. Mevcut veriler değişmez.
            </p>

            <label className="block cursor-pointer">
              <div className="border-2 border-dashed rounded-2xl p-8 text-center active:bg-blue-50 transition-all"
                style={{ borderColor: wfBlue }}>
                <div className="text-4xl mb-2">📂</div>
                <p className="font-bold text-base" style={{ color: wfBlue }}>Excel Dosyası Seç</p>
                <p className="text-gray-400 text-xs mt-1">Stok Kart Kayıtları.xlsx</p>
              </div>
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={stokGuncelle} />
            </label>

            {guncellemeMetni && (
              <div className={`rounded-2xl p-4 text-sm font-medium ${
                guncellemeMetni.startsWith("✅") ? "bg-green-50 text-green-800 border border-green-200" :
                guncellemeMetni.startsWith("❌") ? "bg-red-50 text-red-800 border border-red-200" :
                "bg-gray-50 text-gray-600"}`}>
                {guncellemeMetni}
              </div>
            )}
          </div>

          <button onClick={() => setAdim("scan")}
            className="w-full bg-white border border-gray-200 text-gray-600 font-bold py-4 rounded-3xl text-base shadow-sm">
            ← Geri Dön
          </button>
        </div>
      )}
    </div>
  );
}
