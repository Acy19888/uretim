"use client";

import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { findMatches } from "@/lib/fuzzy";
import type { StokItem, ScannedItem, ConfirmedItem, ScanResult } from "@/lib/types";

import bundledStok from "@/lib/stokData.json";

type Step = "scan" | "processing" | "confirm" | "more_pages" | "done" | "update";

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
  } catch {
    return base;
  }
}

export default function Home() {
  const [adim, setAdim]                     = useState<Step>("scan");
  const [stokData, setStokData]             = useState<StokItem[]>([]);

  const [tumOnaylananlar, setTumOnaylananlar] = useState<ConfirmedItem[]>([]);
  const [isEmriNo, setIsEmriNo]              = useState("");
  const [tarih, setTarih]                    = useState(bugunStr());

  const [ralRenk, setRalRenk]               = useState("");
  const [taramaKonusu, setTaramaKonusu]     = useState<ScanResult | null>(null);
  const [mevcutIndex, setMevcutIndex]        = useState(0);
  const [sayfaOnaylananlar, setSayfaOnaylananlar] = useState<ConfirmedItem[]>([]);
  const [onizlemeGorsel, setOnizlemeGorsel] = useState<string | null>(null);
  const [hata, setHata]                     = useState<string | null>(null);

  const [duzenMiktar, setDuzenMiktar]       = useState("");
  const [duzenStok, setDuzenStok]           = useState<StokItem | null>(null);
  const [stokArama, setStokArama]           = useState("");
  const [stokOneriler, setStokOneriler]     = useState<StokItem[]>([]);
  const [duzenlemeAcik, setDuzenlemeAcik]   = useState(false);

  const [guncellemeMetni, setGuncellemeMetni] = useState<string | null>(null);

  useEffect(() => {
    setStokData(stokYukle());
  }, []);

  // ── Tarama ────────────────────────────────────────────────────────────────
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
      if (!sonuc.items?.length) throw new Error("Hiçbir ürün algılanamadı. Lütfen tekrar deneyin.");
      if (!isEmriNo && sonuc.is_emri_no) setIsEmriNo(sonuc.is_emri_no);
      if (sonuc.tarih) setTarih(sonuc.tarih);
      if (sonuc.ral_renk) setRalRenk(sonuc.ral_renk);
      setTaramaKonusu(sonuc);
      setMevcutIndex(0);
      setSayfaOnaylananlar([]);
      setAdim("confirm");
    } catch (err) { setHata(String(err)); setAdim("scan"); }
  }

  // ── Onay: satırı hazırla ───────────────────────────────────────────────────
  const satirHazirla = useCallback((urun: ScannedItem) => {
    setDuzenMiktar(urun.miktar || "");
    setStokArama(urun.urun_adi || "");
    setDuzenlemeAcik(false);
    // RAL rengi ile birlikte ara: "9005 Monaka Bosma"
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
      original_urun_adi: urun.urun_adi,
      original_miktar: urun.miktar,
      confirmed_stok: duzenStok,
      confirmed_miktar: parseFloat(duzenMiktar.replace(",", ".")) || 0,
      skipped: false,
    }]);
  }

  function atla() {
    const urun = taramaKonusu!.items[mevcutIndex];
    ilerle([...sayfaOnaylananlar, {
      original_urun_adi: urun.urun_adi,
      original_miktar: urun.miktar,
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

  // ── Stok arama ────────────────────────────────────────────────────────────
  function stokAramaGuncelle(q: string) {
    setStokArama(q);
    setStokOneriler(q.length >= 2 ? findMatches(q, stokData, 8) : []);
  }

  // ── Stok güncelle ─────────────────────────────────────────────────────────
  async function stokGuncelle(e: React.ChangeEvent<HTMLInputElement>) {
    const dosya = e.target.files?.[0];
    if (!dosya) return;
    setGuncellemeMetni("Okunuyor...");
    try {
      const buf  = await dosya.arrayBuffer();
      const wb   = XLSX.read(buf, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
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
      const birlestirilen = Array.from(map.values());

      localStorage.setItem("stok_guncellemeler", JSON.stringify(birlestirilen));
      setStokData(stokYukle());
      setGuncellemeMetni(`✅ ${gercektenYeni.length} yeni ürün eklendi. Toplam: ${yeniUrunler.length}`);
    } catch (err) {
      setGuncellemeMetni("❌ Hata: " + String(err));
    }
  }

  // ── Excel dışa aktarma ────────────────────────────────────────────────────
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
      a.href     = url;
      a.download = `Uretim_${isEmriNo || "cikti"}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) { setHata("Dışa aktarma hatası: " + String(err)); }
  }

  const mevcutUrun     = taramaKonusu?.items[mevcutIndex];
  const toplamUrun     = taramaKonusu?.items.length || 0;
  const onaylananSayisi = tumOnaylananlar.filter(i => !i.skipped).length;

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div className="space-y-4 pb-10">

      {/* Hata */}
      {hata && (
        <div className="bg-red-50 border border-red-300 text-red-800 rounded-2xl p-4 flex gap-3 items-start">
          <span className="text-2xl">⚠️</span>
          <div className="flex-1 text-base">{hata}</div>
          <button onClick={() => setHata(null)} className="text-red-400 text-2xl font-bold">✕</button>
        </div>
      )}

      {/* ── TARAMA ───────────────────────────────────────────────────────── */}
      {adim === "scan" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-5 flex justify-between items-center">
            <div>
              <h2 className="text-xl font-bold text-blue-900">📋 Fiş Tara</h2>
              {tumOnaylananlar.length > 0 && (
                <p className="text-sm text-green-600 mt-1">+ {onaylananSayisi} ürün onaylandı</p>
              )}
            </div>
            <div className="text-right">
              <p className="text-xs text-gray-400">Stok Adedi</p>
              <p className="text-lg font-bold text-green-700">{stokData.length.toLocaleString()}</p>
            </div>
          </div>

          {/* Kamera */}
          <label className="block cursor-pointer">
            <div className="bg-blue-700 active:bg-blue-900 text-white rounded-3xl p-10 text-center shadow-lg">
              <div className="text-7xl mb-4">📷</div>
              <p className="text-3xl font-bold">Kamera</p>
              <p className="text-blue-200 mt-2 text-lg">Fotoğraf çek</p>
            </div>
            <input type="file" accept="image/*" capture="environment" className="hidden" onChange={gorselSec} />
          </label>

          {/* Galeriden seç */}
          <label className="block cursor-pointer">
            <div className="bg-white border-2 border-gray-200 rounded-3xl p-8 text-center active:bg-gray-50">
              <div className="text-5xl mb-3">🖼️</div>
              <p className="text-xl font-bold text-gray-700">Galeriden Seç</p>
            </div>
            <input type="file" accept="image/*" className="hidden" onChange={gorselSec} />
          </label>

          {/* Güncelle */}
          <button onClick={() => { setAdim("update"); setGuncellemeMetni(null); }}
            className="w-full bg-gray-100 active:bg-gray-200 text-gray-500 font-medium py-4 rounded-2xl text-base">
            🔄 Stok Kartını Güncelle
          </button>
        </div>
      )}

      {/* ── İŞLENİYOR ────────────────────────────────────────────────────── */}
      {adim === "processing" && (
        <div className="bg-white rounded-3xl shadow-md p-10 text-center space-y-6">
          {onizlemeGorsel && (
            // eslint-disable-next-line @next/next/no-img-element
            <img src={onizlemeGorsel} alt="" className="max-h-56 mx-auto rounded-2xl shadow object-contain" />
          )}
          <div className="text-6xl animate-spin">⚙️</div>
          <div>
            <p className="text-2xl font-bold text-blue-900">Analiz ediliyor...</p>
            <p className="text-gray-400 mt-1">Yapay zeka el yazısını okuyor</p>
          </div>
        </div>
      )}

      {/* ── ONAY ─────────────────────────────────────────────────────────── */}
      {adim === "confirm" && mevcutUrun && (
        <div className="space-y-4">
          {/* İlerleme */}
          <div className="bg-white rounded-3xl shadow-md p-4">
            <div className="flex justify-between text-sm text-gray-500 mb-2">
              <span className="font-bold text-gray-700">{mevcutIndex + 1}. Satır / {toplamUrun}</span>
              <span>{ralRenk && <span className="bg-gray-800 text-white px-2 py-0.5 rounded-lg text-xs font-mono mr-2">RAL {ralRenk}</span>}{isEmriNo && `İş Emri: ${isEmriNo}`}</span>
            </div>
            <div className="w-full bg-gray-100 rounded-full h-3">
              <div className="bg-blue-700 h-3 rounded-full transition-all"
                style={{ width: `${(mevcutIndex / toplamUrun) * 100}%` }} />
            </div>
          </div>

          {/* Fiş içeriği */}
          <div className="bg-amber-50 border-2 border-amber-200 rounded-3xl p-5 space-y-3">
            <p className="text-xs font-bold text-amber-600 uppercase tracking-widest">📄 Fişten Okunan</p>
            <p className="text-2xl font-bold text-gray-900 leading-tight">{mevcutUrun.urun_adi}</p>
            <p className="text-5xl font-black text-blue-900">{mevcutUrun.miktar}</p>
          </div>

          {/* Stok eşleşmesi */}
          {!duzenlemeAcik ? (
            <div className="bg-white rounded-3xl shadow-md p-5 space-y-4">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">🔍 Stok Eşleşmesi</p>
              {duzenStok ? (
                <div className="bg-green-50 border-2 border-green-200 rounded-2xl p-4">
                  <p className="font-mono text-green-700 font-bold text-lg">{duzenStok.stok_kodu}</p>
                  <p className="font-bold text-gray-900 text-lg mt-1 leading-snug">{duzenStok.stok_adi}</p>
                </div>
              ) : (
                <div className="bg-red-50 border-2 border-red-200 rounded-2xl p-4 text-center">
                  <p className="text-red-600 font-bold text-lg">Eşleşen ürün bulunamadı</p>
                </div>
              )}

              {stokOneriler.length > 1 && (
                <div className="space-y-2">
                  <p className="text-sm text-gray-400">Diğer öneriler:</p>
                  {stokOneriler.slice(1, 4).map(s => (
                    <button key={s.stok_kodu} onClick={() => setDuzenStok(s)}
                      className="w-full text-left text-sm bg-gray-50 active:bg-blue-50 border border-gray-200 rounded-xl px-4 py-3">
                      <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span><br />{s.stok_adi}
                    </button>
                  ))}
                </div>
              )}

              <div className="flex items-center justify-between pt-2 border-t border-gray-100">
                <div>
                  <p className="text-sm text-gray-400">Miktar</p>
                  <p className="text-4xl font-black text-blue-900">{duzenMiktar}</p>
                </div>
                <button onClick={() => setDuzenlemeAcik(true)}
                  className="bg-gray-100 active:bg-gray-200 text-gray-700 font-bold px-5 py-3 rounded-2xl text-base">
                  ✏️ Düzenle
                </button>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-3xl shadow-md p-5 space-y-4">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">✏️ Düzenleme</p>
              <div>
                <label className="text-sm font-bold text-gray-600 block mb-2">Miktar</label>
                <input type="number" value={duzenMiktar} onChange={e => setDuzenMiktar(e.target.value)} autoFocus
                  className="w-full border-2 border-blue-400 rounded-2xl p-4 text-3xl font-bold text-center focus:outline-none focus:border-blue-700" />
              </div>
              <div>
                <label className="text-sm font-bold text-gray-600 block mb-2">Ürün Ara</label>
                <input type="text" value={stokArama} onChange={e => stokAramaGuncelle(e.target.value)}
                  placeholder="Ürün adı veya stok kodu..."
                  className="w-full border-2 border-gray-300 rounded-2xl p-4 text-base focus:outline-none focus:border-blue-500" />
                {stokOneriler.length > 0 && (
                  <div className="mt-2 space-y-1 max-h-52 overflow-y-auto">
                    {stokOneriler.map(s => (
                      <button key={s.stok_kodu}
                        onClick={() => { setDuzenStok(s); setStokArama(s.stok_adi); setStokOneriler([]); }}
                        className={`w-full text-left text-sm rounded-xl px-4 py-3 border transition-colors ${duzenStok?.stok_kodu === s.stok_kodu ? "bg-green-50 border-green-300 font-bold" : "bg-gray-50 border-gray-200 active:bg-blue-50"}`}>
                        <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span><br />{s.stok_adi}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <button onClick={() => setDuzenlemeAcik(false)}
                className="w-full bg-blue-700 active:bg-blue-900 text-white font-bold py-5 rounded-2xl text-xl">
                💾 Kaydet
              </button>
            </div>
          )}

          {!duzenlemeAcik && (
            <div className="grid grid-cols-2 gap-3">
              <button onClick={atla}
                className="bg-gray-200 active:bg-gray-300 text-gray-800 font-bold py-6 rounded-3xl text-xl">
                ⏭ Atla
              </button>
              <button onClick={onayla}
                className="bg-green-600 active:bg-green-800 text-white font-bold py-6 rounded-3xl text-2xl shadow-lg">
                ✅ Evet
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── BAŞKA SAYFA ──────────────────────────────────────────────────── */}
      {adim === "more_pages" && (
        <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-6">
          <div className="text-6xl">📄</div>
          <h2 className="text-2xl font-bold text-gray-900">Başka Sayfa Var Mı?</h2>
          <p className="text-gray-500 text-lg">
            {tumOnaylananlar.filter(i => !i.skipped).length} ürün onaylandı
          </p>
          <div className="grid grid-cols-2 gap-4">
            <button onClick={() => {
              setTaramaKonusu(null); setSayfaOnaylananlar([]);
              setMevcutIndex(0); setOnizlemeGorsel(null); setAdim("scan");
            }} className="bg-green-600 active:bg-green-800 text-white font-bold py-8 rounded-3xl text-2xl shadow-lg">
              ✅ Evet
            </button>
            <button onClick={() => setAdim("done")}
              className="bg-red-600 active:bg-red-800 text-white font-bold py-8 rounded-3xl text-2xl shadow-lg">
              ❌ Hayır
            </button>
          </div>
          <p className="text-gray-400 text-sm">Hayır → Excel oluşturulur</p>
        </div>
      )}

      {/* ── TAMAMLANDI ───────────────────────────────────────────────────── */}
      {adim === "done" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-4">
            <div className="text-7xl">✅</div>
            <h2 className="text-3xl font-bold text-green-700">Tamamlandı!</h2>
            <p className="text-xl text-gray-600">
              <strong>{onaylananSayisi}</strong> ürün — İş Emri: {isEmriNo || "—"}
            </p>
          </div>

          <div className="bg-white rounded-3xl shadow-md p-5 space-y-2">
            <h3 className="font-bold text-gray-700 text-lg">Onaylanan Ürünler</h3>
            <div className="space-y-2 max-h-64 overflow-y-auto">
              {tumOnaylananlar.map((urun, idx) => (
                <div key={idx} className={`flex items-center gap-3 p-3 rounded-2xl ${urun.skipped ? "bg-gray-50 text-gray-400" : "bg-green-50"}`}>
                  <span className="text-xl">{urun.skipped ? "⏭" : "✅"}</span>
                  <div className="flex-1 min-w-0">
                    <p className="font-medium truncate text-base">{urun.confirmed_stok?.stok_adi || urun.original_urun_adi}</p>
                    {urun.confirmed_stok && <p className="text-xs font-mono text-gray-400">{urun.confirmed_stok.stok_kodu}</p>}
                  </div>
                  {!urun.skipped && <span className="font-bold text-blue-900 text-xl whitespace-nowrap">{urun.confirmed_miktar}</span>}
                </div>
              ))}
            </div>
          </div>

          <button onClick={excelIndir}
            className="w-full bg-blue-700 active:bg-blue-900 text-white font-bold py-7 rounded-3xl text-2xl shadow-xl">
            📥 Excel İndir
          </button>

          <button onClick={() => {
            setTumOnaylananlar([]); setTaramaKonusu(null); setSayfaOnaylananlar([]);
            setMevcutIndex(0); setOnizlemeGorsel(null); setIsEmriNo(""); setAdim("scan");
          }} className="w-full bg-gray-200 active:bg-gray-300 text-gray-700 font-bold py-5 rounded-3xl text-xl">
            🔄 Yeni Fiş Tara
          </button>
        </div>
      )}

      {/* ── GÜNCELLE ─────────────────────────────────────────────────────── */}
      {adim === "update" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-6 space-y-4">
            <h2 className="text-2xl font-bold text-blue-900">🔄 Stok Kartını Güncelle</h2>
            <p className="text-gray-500">
              Yeni stok Excel dosyasını yükleyin. Yalnızca <strong>yeni stok kodları</strong> eklenir. Mevcut kodlar değiştirilmez.
            </p>
            <div className="bg-blue-50 rounded-2xl p-4 text-sm text-blue-800">
              Sistemde kayıtlı: <strong>{stokData.length.toLocaleString()}</strong> ürün
            </div>

            <label className="block cursor-pointer">
              <div className="border-2 border-dashed border-blue-400 rounded-2xl p-8 text-center active:bg-blue-50">
                <div className="text-5xl mb-3">📂</div>
                <p className="text-blue-700 font-bold text-xl">Excel Dosyası Seç</p>
                <p className="text-gray-400 mt-1">Yeni Stok Kart Kayıtları.xlsx</p>
              </div>
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={stokGuncelle} />
            </label>

            {guncellemeMetni && (
              <div className={`rounded-2xl p-4 text-base font-medium ${
                guncellemeMetni.startsWith("✅") ? "bg-green-50 text-green-800" :
                guncellemeMetni.startsWith("❌") ? "bg-red-50 text-red-800" :
                "bg-gray-50 text-gray-600"}`}>
                {guncellemeMetni}
              </div>
            )}
          </div>

          <button onClick={() => setAdim("scan")}
            className="w-full bg-gray-200 active:bg-gray-300 text-gray-800 font-bold py-5 rounded-3xl text-xl">
            ← Geri Dön
          </button>
        </div>
      )}
    </div>
  );
}
