"use client";

import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { findMatches } from "@/lib/fuzzy";
import type { StokItem, ScannedItem, ConfirmedItem, ScanResult } from "@/lib/types";

type Step = "setup" | "scan" | "processing" | "confirm" | "more_pages" | "done";

function toBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve((reader.result as string).split(",")[1]);
    reader.onerror = reject;
  });
}

function todayStr(): string {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
}

export default function Home() {
  const [step, setStep] = useState<Step>("setup");
  const [stokData, setStokData] = useState<StokItem[]>([]);

  // Accumulates items across multiple pages
  const [allConfirmedItems, setAllConfirmedItems] = useState<ConfirmedItem[]>([]);
  const [isEmriNo, setIsEmriNo] = useState("");
  const [tarih, setTarih] = useState(todayStr());

  // Per-page state
  const [scanResult, setScanResult] = useState<ScanResult | null>(null);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [pageConfirmed, setPageConfirmed] = useState<ConfirmedItem[]>([]);
  const [previewImage, setPreviewImage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Edit state
  const [editMiktar, setEditMiktar] = useState("");
  const [editStok, setEditStok] = useState<StokItem | null>(null);
  const [stokSearch, setStokSearch] = useState("");
  const [stokSuggestions, setStokSuggestions] = useState<StokItem[]>([]);
  const [isEditing, setIsEditing] = useState(false);

  // Load stok from localStorage
  useEffect(() => {
    const stored = localStorage.getItem("stok_data");
    if (stored) {
      try {
        const parsed: StokItem[] = JSON.parse(stored);
        if (parsed.length > 0) { setStokData(parsed); setStep("scan"); }
      } catch { localStorage.removeItem("stok_data"); }
    }
  }, []);

  // ── Setup ──────────────────────────────────────────────────────────────────
  async function handleStokUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, string>>(ws, { raw: false, defval: "" });
      const items: StokItem[] = rows.flatMap(row => {
        const kodu = row["STOK_KODU"] || row["Stok Kodu"] || row["STOK KODU"] || "";
        const adi  = row["STOK_ADI"]  || row["Stok Adı"]  || row["STOK ADI"]  || "";
        if (!kodu || !adi) return [];
        return [{ stok_kodu: String(kodu).trim(), stok_adi: String(adi).trim(), cesit: String(row["Çeşit"] || "").trim() }];
      });
      if (items.length === 0) { setError("Spalten STOK_KODU und STOK_ADI nicht gefunden."); return; }
      localStorage.setItem("stok_data", JSON.stringify(items));
      setStokData(items);
      setStep("scan");
    } catch (err) { setError("Excel-Lesefehler: " + String(err)); }
  }

  // ── Scan ───────────────────────────────────────────────────────────────────
  async function handleImageSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(null);
    setPreviewImage(URL.createObjectURL(file));
    setStep("processing");
    try {
      const base64 = await toBase64(file);
      const res = await fetch("/api/scan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ imageBase64: base64, mimeType: file.type }),
      });
      if (!res.ok) throw new Error((await res.json()).error || "Scan fehlgeschlagen");
      const result: ScanResult = await res.json();
      if (!result.items?.length) throw new Error("Keine Produkte erkannt.");
      // Set İş Emri from first page if not already set
      if (!isEmriNo && result.is_emri_no) setIsEmriNo(result.is_emri_no);
      if (result.tarih) setTarih(result.tarih);
      setScanResult(result);
      setCurrentIndex(0);
      setPageConfirmed([]);
      setStep("confirm");
    } catch (err) { setError(String(err)); setStep("scan"); }
  }

  // ── Confirm: prepare item ──────────────────────────────────────────────────
  const prepareItem = useCallback((item: ScannedItem) => {
    setEditMiktar(item.miktar || "");
    setStokSearch(item.urun_adi || "");
    setIsEditing(false);
    if (stokData.length > 0) {
      const matches = findMatches(item.urun_adi, stokData, 5);
      setStokSuggestions(matches);
      setEditStok(matches[0] || null);
    }
  }, [stokData]);

  useEffect(() => {
    if (step === "confirm" && scanResult && currentIndex < scanResult.items.length)
      prepareItem(scanResult.items[currentIndex]);
  }, [currentIndex, step, scanResult, prepareItem]);

  function handleConfirm() {
    const item = scanResult!.items[currentIndex];
    const qty = parseFloat(editMiktar.replace(",", ".")) || 0;
    advance([...pageConfirmed, {
      original_urun_adi: item.urun_adi,
      original_miktar: item.miktar,
      confirmed_stok: editStok,
      confirmed_miktar: qty,
      skipped: false,
    }]);
  }

  function handleSkip() {
    const item = scanResult!.items[currentIndex];
    advance([...pageConfirmed, {
      original_urun_adi: item.urun_adi,
      original_miktar: item.miktar,
      confirmed_stok: null,
      confirmed_miktar: 0,
      skipped: true,
    }]);
  }

  function advance(updated: ConfirmedItem[]) {
    setPageConfirmed(updated);
    if (currentIndex + 1 < scanResult!.items.length) {
      setCurrentIndex(currentIndex + 1);
    } else {
      // Page done → ask for more pages
      setAllConfirmedItems(prev => [...prev, ...updated]);
      setStep("more_pages");
    }
  }

  // ── More pages ─────────────────────────────────────────────────────────────
  function handleMorePagesYes() {
    setScanResult(null);
    setPageConfirmed([]);
    setCurrentIndex(0);
    setPreviewImage(null);
    setStep("scan");
  }

  function handleMorePagesNo() {
    setStep("done");
  }

  // ── Stok search ────────────────────────────────────────────────────────────
  function handleStokSearch(q: string) {
    setStokSearch(q);
    setStokSuggestions(q.length >= 2 ? findMatches(q, stokData, 8) : []);
  }

  // ── Export ─────────────────────────────────────────────────────────────────
  async function handleExport() {
    setError(null);
    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ is_emri_no: isEmriNo, tarih, items: allConfirmedItems }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `Uretim_${isEmriNo || "export"}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) { setError("Export fehlgeschlagen: " + String(err)); }
  }

  const currentItem = scanResult?.items[currentIndex];
  const totalItems = scanResult?.items.length || 0;
  const confirmedCount = allConfirmedItems.filter(i => !i.skipped).length + pageConfirmed.filter(i => !i.skipped).length;

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div className="space-y-4 pb-8">

      {/* Error */}
      {error && (
        <div className="bg-red-50 border border-red-300 text-red-800 rounded-2xl p-4 flex gap-3 items-start">
          <span className="text-2xl mt-0.5">⚠️</span>
          <div className="flex-1 text-base">{error}</div>
          <button onClick={() => setError(null)} className="text-red-400 text-xl font-bold">✕</button>
        </div>
      )}

      {/* ── SETUP ──────────────────────────────────────────────────────────── */}
      {step === "setup" && (
        <div className="bg-white rounded-3xl shadow-md p-6 space-y-5">
          <h2 className="text-2xl font-bold text-blue-900">🔧 Stok Verisi Yükle</h2>
          <p className="text-gray-500 text-base">
            Einmalig <strong>Stok Kart Kayıtları.xlsx</strong> hochladen. Wird gespeichert.
          </p>
          <label className="block cursor-pointer">
            <div className="border-2 border-dashed border-blue-400 rounded-2xl p-10 text-center active:bg-blue-50">
              <div className="text-5xl mb-3">📂</div>
              <p className="text-blue-700 font-bold text-xl">Excel wählen</p>
              <p className="text-gray-400 mt-1">Stok Kart Kayıtları.xlsx</p>
            </div>
            <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleStokUpload} />
          </label>
        </div>
      )}

      {/* ── SCAN ───────────────────────────────────────────────────────────── */}
      {step === "scan" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-5 flex justify-between items-center">
            <div>
              <h2 className="text-xl font-bold text-blue-900">📷 Zettel Scannen</h2>
              {allConfirmedItems.length > 0 && (
                <p className="text-sm text-green-600 mt-1">
                  + {allConfirmedItems.filter(i=>!i.skipped).length} bereits bestätigt
                </p>
              )}
            </div>
            <span className="text-sm text-green-600 font-medium bg-green-50 px-3 py-1 rounded-full whitespace-nowrap">
              ✅ {stokData.length.toLocaleString()}
            </span>
          </div>

          <label className="block cursor-pointer">
            <div className="bg-blue-700 active:bg-blue-900 text-white rounded-3xl p-10 text-center shadow-lg transition-all">
              <div className="text-7xl mb-4">📷</div>
              <p className="text-3xl font-bold">Kamera</p>
              <p className="text-blue-200 mt-2 text-lg">Foto aufnehmen</p>
            </div>
            <input type="file" accept="image/*" capture="environment" className="hidden" onChange={handleImageSelect} />
          </label>

          <label className="block cursor-pointer">
            <div className="bg-white border-2 border-gray-200 rounded-3xl p-8 text-center active:bg-gray-50">
              <div className="text-5xl mb-3">🖼️</div>
              <p className="text-xl font-bold text-gray-700">Aus Galerie</p>
            </div>
            <input type="file" accept="image/*" className="hidden" onChange={handleImageSelect} />
          </label>

          <button onClick={() => setStep("setup")} className="w-full text-center text-gray-400 text-sm underline py-2">
            Stok Datei wechseln
          </button>
        </div>
      )}

      {/* ── PROCESSING ─────────────────────────────────────────────────────── */}
      {step === "processing" && (
        <div className="bg-white rounded-3xl shadow-md p-10 text-center space-y-6">
          {previewImage && (
            // eslint-disable-next-line @next/next/no-img-element
            <img src={previewImage} alt="" className="max-h-56 mx-auto rounded-2xl shadow object-contain" />
          )}
          <div className="text-6xl animate-spin">⚙️</div>
          <div>
            <p className="text-2xl font-bold text-blue-900">Wird analysiert...</p>
            <p className="text-gray-400 mt-1">KI liest die Handschrift</p>
          </div>
        </div>
      )}

      {/* ── CONFIRM ────────────────────────────────────────────────────────── */}
      {step === "confirm" && currentItem && (
        <div className="space-y-4">
          {/* Progress bar */}
          <div className="bg-white rounded-3xl shadow-md p-4">
            <div className="flex justify-between text-sm text-gray-500 mb-2">
              <span className="font-bold text-gray-700">Zeile {currentIndex + 1} / {totalItems}</span>
              <span>{isEmriNo && `İş Emri: ${isEmriNo}`}</span>
            </div>
            <div className="w-full bg-gray-100 rounded-full h-3">
              <div
                className="bg-blue-700 h-3 rounded-full transition-all"
                style={{ width: `${(currentIndex / totalItems) * 100}%` }}
              />
            </div>
          </div>

          {/* OCR result */}
          <div className="bg-amber-50 border-2 border-amber-200 rounded-3xl p-5 space-y-3">
            <p className="text-xs font-bold text-amber-600 uppercase tracking-widest">📖 Zettel</p>
            <p className="text-2xl font-bold text-gray-900 leading-tight">{currentItem.urun_adi}</p>
            <p className="text-5xl font-black text-blue-900">{currentItem.miktar}</p>
          </div>

          {/* Matched product */}
          {!isEditing ? (
            <div className="bg-white rounded-3xl shadow-md p-5 space-y-4">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">🔍 Stok Eşleşme</p>
              {editStok ? (
                <div className="bg-green-50 border-2 border-green-200 rounded-2xl p-4">
                  <p className="font-mono text-green-700 font-bold text-lg">{editStok.stok_kodu}</p>
                  <p className="font-bold text-gray-900 text-lg mt-1 leading-snug">{editStok.stok_adi}</p>
                </div>
              ) : (
                <div className="bg-red-50 border-2 border-red-200 rounded-2xl p-4 text-center">
                  <p className="text-red-600 font-bold text-lg">Kein Produkt gefunden</p>
                </div>
              )}

              {/* Other suggestions */}
              {stokSuggestions.length > 1 && (
                <div className="space-y-2">
                  <p className="text-sm text-gray-400">Andere Vorschläge:</p>
                  {stokSuggestions.slice(1, 4).map(s => (
                    <button key={s.stok_kodu} onClick={() => setEditStok(s)}
                      className="w-full text-left text-sm bg-gray-50 active:bg-blue-50 border border-gray-200 rounded-xl px-4 py-3 transition-colors">
                      <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span>
                      <br />{s.stok_adi}
                    </button>
                  ))}
                </div>
              )}

              {/* Quantity + edit */}
              <div className="flex items-center justify-between pt-2 border-t border-gray-100">
                <div>
                  <p className="text-sm text-gray-400">Menge</p>
                  <p className="text-4xl font-black text-blue-900">{editMiktar}</p>
                </div>
                <button onClick={() => setIsEditing(true)}
                  className="bg-gray-100 active:bg-gray-200 text-gray-700 font-bold px-5 py-3 rounded-2xl text-base">
                  ✏️ Bearbeiten
                </button>
              </div>
            </div>
          ) : (
            /* Edit mode */
            <div className="bg-white rounded-3xl shadow-md p-5 space-y-4">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">✏️ Bearbeiten</p>
              <div>
                <label className="text-sm font-bold text-gray-600 block mb-2">Menge</label>
                <input type="number" value={editMiktar} onChange={e => setEditMiktar(e.target.value)} autoFocus
                  className="w-full border-2 border-blue-400 rounded-2xl p-4 text-3xl font-bold text-center focus:outline-none focus:border-blue-700" />
              </div>
              <div>
                <label className="text-sm font-bold text-gray-600 block mb-2">Produkt suchen</label>
                <input type="text" value={stokSearch} onChange={e => handleStokSearch(e.target.value)}
                  placeholder="Name oder Stok Kodu..."
                  className="w-full border-2 border-gray-300 rounded-2xl p-4 text-base focus:outline-none focus:border-blue-500" />
                {stokSuggestions.length > 0 && (
                  <div className="mt-2 space-y-1 max-h-52 overflow-y-auto">
                    {stokSuggestions.map(s => (
                      <button key={s.stok_kodu} onClick={() => { setEditStok(s); setStokSearch(s.stok_adi); setStokSuggestions([]); }}
                        className={`w-full text-left text-sm rounded-xl px-4 py-3 border transition-colors ${editStok?.stok_kodu === s.stok_kodu ? "bg-green-50 border-green-300 font-bold" : "bg-gray-50 border-gray-200 active:bg-blue-50"}`}>
                        <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span><br />{s.stok_adi}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <button onClick={() => setIsEditing(false)}
                className="w-full bg-blue-700 active:bg-blue-900 text-white font-bold py-5 rounded-2xl text-xl">
                💾 Speichern
              </button>
            </div>
          )}

          {/* Action buttons */}
          {!isEditing && (
            <div className="grid grid-cols-2 gap-3">
              <button onClick={handleSkip}
                className="bg-gray-200 active:bg-gray-300 text-gray-800 font-bold py-6 rounded-3xl text-xl transition-colors">
                ⏭ Überspringen
              </button>
              <button onClick={handleConfirm}
                className="bg-green-600 active:bg-green-800 text-white font-bold py-6 rounded-3xl text-xl shadow-lg transition-colors">
                ✅ Ja
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── MORE PAGES ─────────────────────────────────────────────────────── */}
      {step === "more_pages" && (
        <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-6">
          <div className="text-6xl">📄</div>
          <div>
            <h2 className="text-2xl font-bold text-gray-900">Weitere Seite?</h2>
            <p className="text-gray-500 mt-2 text-lg">
              {allConfirmedItems.filter(i => !i.skipped).length} Produkte bestätigt
            </p>
          </div>
          <div className="grid grid-cols-2 gap-4 pt-2">
            <button onClick={handleMorePagesNo}
              className="bg-red-600 active:bg-red-800 text-white font-bold py-8 rounded-3xl text-2xl shadow-lg">
              ❌ Nein
            </button>
            <button onClick={handleMorePagesYes}
              className="bg-green-600 active:bg-green-800 text-white font-bold py-8 rounded-3xl text-2xl shadow-lg">
              ✅ Ja
            </button>
          </div>
          <p className="text-gray-400 text-sm">Nein = Excel wird erstellt</p>
        </div>
      )}

      {/* ── DONE ───────────────────────────────────────────────────────────── */}
      {step === "done" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-4">
            <div className="text-7xl">✅</div>
            <h2 className="text-3xl font-bold text-green-700">Fertig!</h2>
            <p className="text-xl text-gray-600">
              <strong>{confirmedCount}</strong> Produkte — İş Emri {isEmriNo || "—"}
            </p>
          </div>

          {/* Summary */}
          <div className="bg-white rounded-3xl shadow-md p-5 space-y-3">
            <h3 className="font-bold text-gray-700 text-lg">Übersicht</h3>
            <div className="space-y-2 max-h-64 overflow-y-auto">
              {allConfirmedItems.map((item, idx) => (
                <div key={idx} className={`flex items-center gap-3 p-3 rounded-2xl text-sm ${item.skipped ? "bg-gray-50 text-gray-400" : "bg-green-50"}`}>
                  <span className="text-xl">{item.skipped ? "⏭" : "✅"}</span>
                  <div className="flex-1 min-w-0">
                    <p className="font-medium truncate text-base">
                      {item.confirmed_stok?.stok_adi || item.original_urun_adi}
                    </p>
                    {item.confirmed_stok && (
                      <p className="text-xs font-mono text-gray-400">{item.confirmed_stok.stok_kodu}</p>
                    )}
                  </div>
                  {!item.skipped && (
                    <span className="font-bold text-blue-900 text-xl whitespace-nowrap">{item.confirmed_miktar}</span>
                  )}
                </div>
              ))}
            </div>
          </div>

          <button onClick={handleExport}
            className="w-full bg-blue-700 active:bg-blue-900 text-white font-bold py-7 rounded-3xl text-2xl shadow-xl">
            📥 Excel Herunterladen
          </button>

          <button onClick={() => {
            setAllConfirmedItems([]); setScanResult(null); setPageConfirmed([]);
            setCurrentIndex(0); setPreviewImage(null); setIsEmriNo(""); setStep("scan");
          }} className="w-full bg-gray-200 active:bg-gray-300 text-gray-700 font-bold py-5 rounded-3xl text-xl">
            🔄 Neu starten
          </button>
        </div>
      )}
    </div>
  );
}
