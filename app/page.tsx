"use client";

import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { findMatches } from "@/lib/fuzzy";
import type { StokItem, ScannedItem, ConfirmedItem, ScanResult } from "@/lib/types";

// Bundled hard-coded data (generated via: node scripts/convert-stok.js)
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

function todayStr(): string {
  const d = new Date();
  return `${String(d.getDate()).padStart(2,"0")}.${String(d.getMonth()+1).padStart(2,"0")}.${d.getFullYear()}`;
}

function loadStok(): StokItem[] {
  // Merge: bundled base + localStorage updates
  const base = bundledStok as StokItem[];
  try {
    const raw = localStorage.getItem("stok_updates");
    if (!raw) return base;
    const updates: StokItem[] = JSON.parse(raw);
    // Updates override base if same kodu, else append
    const baseMap = new Map(base.map(i => [i.stok_kodu, i]));
    for (const u of updates) baseMap.set(u.stok_kodu, u);
    return Array.from(baseMap.values());
  } catch {
    return base;
  }
}

export default function Home() {
  const [step, setStep]               = useState<Step>("scan");
  const [stokData, setStokData]       = useState<StokItem[]>([]);

  // Accumulated across multiple pages
  const [allConfirmed, setAllConfirmed] = useState<ConfirmedItem[]>([]);
  const [isEmriNo, setIsEmriNo]         = useState("");
  const [tarih, setTarih]               = useState(todayStr());

  // Per-page
  const [scanResult, setScanResult]     = useState<ScanResult | null>(null);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [pageConfirmed, setPageConfirmed] = useState<ConfirmedItem[]>([]);
  const [previewImage, setPreviewImage] = useState<string | null>(null);
  const [error, setError]               = useState<string | null>(null);

  // Edit state
  const [editMiktar, setEditMiktar]     = useState("");
  const [editStok, setEditStok]         = useState<StokItem | null>(null);
  const [stokSearch, setStokSearch]     = useState("");
  const [stokSugg, setStokSugg]         = useState<StokItem[]>([]);
  const [isEditing, setIsEditing]       = useState(false);

  // Update screen state
  const [updateStatus, setUpdateStatus] = useState<string | null>(null);

  useEffect(() => {
    setStokData(loadStok());
  }, []);

  // ── Scan ───────────────────────────────────────────────────────────────────
  async function handleImageSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(null);
    setPreviewImage(URL.createObjectURL(file));
    setStep("processing");
    try {
      const b64 = await toBase64(file);
      const res = await fetch("/api/scan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ imageBase64: b64, mimeType: file.type }),
      });
      if (!res.ok) throw new Error((await res.json()).error || "Scan fehlgeschlagen");
      const result: ScanResult = await res.json();
      if (!result.items?.length) throw new Error("Keine Produkte erkannt.");
      if (!isEmriNo && result.is_emri_no) setIsEmriNo(result.is_emri_no);
      if (result.tarih) setTarih(result.tarih);
      setScanResult(result);
      setCurrentIndex(0);
      setPageConfirmed([]);
      setStep("confirm");
    } catch (err) { setError(String(err)); setStep("scan"); }
  }

  // ── Confirm: prepare ───────────────────────────────────────────────────────
  const prepareItem = useCallback((item: ScannedItem) => {
    setEditMiktar(item.miktar || "");
    setStokSearch(item.urun_adi || "");
    setIsEditing(false);
    const matches = findMatches(item.urun_adi, stokData, 5);
    setStokSugg(matches);
    setEditStok(matches[0] || null);
  }, [stokData]);

  useEffect(() => {
    if (step === "confirm" && scanResult && currentIndex < scanResult.items.length)
      prepareItem(scanResult.items[currentIndex]);
  }, [currentIndex, step, scanResult, prepareItem]);

  function handleConfirm() {
    const item = scanResult!.items[currentIndex];
    advance([...pageConfirmed, {
      original_urun_adi: item.urun_adi,
      original_miktar: item.miktar,
      confirmed_stok: editStok,
      confirmed_miktar: parseFloat(editMiktar.replace(",", ".")) || 0,
      skipped: false,
    }]);
  }

  function handleSkip() {
    const item = scanResult!.items[currentIndex];
    advance([...pageConfirmed, {
      original_urun_adi: item.urun_adi,
      original_miktar: item.miktar,
      confirmed_stok: null, confirmed_miktar: 0, skipped: true,
    }]);
  }

  function advance(updated: ConfirmedItem[]) {
    setPageConfirmed(updated);
    if (currentIndex + 1 < scanResult!.items.length) {
      setCurrentIndex(currentIndex + 1);
    } else {
      setAllConfirmed(prev => [...prev, ...updated]);
      setStep("more_pages");
    }
  }

  // ── Stok search ────────────────────────────────────────────────────────────
  function handleStokSearch(q: string) {
    setStokSearch(q);
    setStokSugg(q.length >= 2 ? findMatches(q, stokData, 8) : []);
  }

  // ── Update: merge new Excel ────────────────────────────────────────────────
  async function handleUpdateExcel(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setUpdateStatus("Wird gelesen...");
    try {
      const buf  = await file.arrayBuffer();
      const wb   = XLSX.read(buf, { type: "array" });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string,string>>(ws, { raw: false, defval: "" });

      const newItems: StokItem[] = rows.flatMap(row => {
        const kodu  = (row["STOK_KODU"] || row["Stok Kodu"] || "").trim();
        const adi   = (row["STOK_ADI"]  || row["Stok Adı"]  || "").trim();
        const cesit = (row["Çeşit"] || "").trim();
        if (!kodu || !adi) return [];
        return [{ stok_kodu: kodu, stok_adi: adi, cesit }];
      });

      // Find truly new items (not in bundled base)
      const bundledKodular = new Set((bundledStok as StokItem[]).map(i => i.stok_kodu));
      const reallyNew = newItems.filter(i => !bundledKodular.has(i.stok_kodu));

      // Merge with existing updates
      const existingRaw = localStorage.getItem("stok_updates");
      const existing: StokItem[] = existingRaw ? JSON.parse(existingRaw) : [];
      const existingMap = new Map(existing.map(i => [i.stok_kodu, i]));
      for (const n of reallyNew) existingMap.set(n.stok_kodu, n);
      const merged = Array.from(existingMap.values());

      localStorage.setItem("stok_updates", JSON.stringify(merged));
      setStokData(loadStok());
      setUpdateStatus(`✅ ${reallyNew.length} neue Produkte hinzugefügt. Gesamt: ${newItems.length}`);
    } catch (err) {
      setUpdateStatus("❌ Fehler: " + String(err));
    }
  }

  // ── Export ─────────────────────────────────────────────────────────────────
  async function handleExport() {
    setError(null);
    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ is_emri_no: isEmriNo, tarih, items: allConfirmed }),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      a.download = `Uretim_${isEmriNo || "export"}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) { setError("Export: " + String(err)); }
  }

  const currentItem   = scanResult?.items[currentIndex];
  const totalItems    = scanResult?.items.length || 0;
  const confirmedCount = allConfirmed.filter(i => !i.skipped).length;

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div className="space-y-4 pb-10">

      {/* Error */}
      {error && (
        <div className="bg-red-50 border border-red-300 text-red-800 rounded-2xl p-4 flex gap-3 items-start">
          <span className="text-2xl">⚠️</span>
          <div className="flex-1 text-base">{error}</div>
          <button onClick={() => setError(null)} className="text-red-400 text-2xl font-bold">✕</button>
        </div>
      )}

      {/* ── SCAN ─────────────────────────────────────────────────────────── */}
      {step === "scan" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-5 flex justify-between items-center">
            <div>
              <h2 className="text-xl font-bold text-blue-900">📷 Zettel Scannen</h2>
              {allConfirmed.length > 0 && (
                <p className="text-sm text-green-600 mt-1">+ {confirmedCount} bereits bestätigt</p>
              )}
            </div>
            <div className="text-right">
              <p className="text-xs text-gray-400">Produkte</p>
              <p className="text-lg font-bold text-green-700">{stokData.length.toLocaleString()}</p>
            </div>
          </div>

          {/* Camera */}
          <label className="block cursor-pointer">
            <div className="bg-blue-700 active:bg-blue-900 text-white rounded-3xl p-10 text-center shadow-lg">
              <div className="text-7xl mb-4">📷</div>
              <p className="text-3xl font-bold">Kamera</p>
              <p className="text-blue-200 mt-2 text-lg">Foto aufnehmen</p>
            </div>
            <input type="file" accept="image/*" capture="environment" className="hidden" onChange={handleImageSelect} />
          </label>

          {/* Gallery */}
          <label className="block cursor-pointer">
            <div className="bg-white border-2 border-gray-200 rounded-3xl p-8 text-center active:bg-gray-50">
              <div className="text-5xl mb-3">🖼️</div>
              <p className="text-xl font-bold text-gray-700">Aus Galerie</p>
            </div>
            <input type="file" accept="image/*" className="hidden" onChange={handleImageSelect} />
          </label>

          {/* Update button */}
          <button onClick={() => { setStep("update"); setUpdateStatus(null); }}
            className="w-full bg-gray-100 active:bg-gray-200 text-gray-500 font-medium py-4 rounded-2xl text-base">
            🔄 Stok aktualisieren
          </button>
        </div>
      )}

      {/* ── PROCESSING ───────────────────────────────────────────────────── */}
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

      {/* ── CONFIRM ──────────────────────────────────────────────────────── */}
      {step === "confirm" && currentItem && (
        <div className="space-y-4">
          {/* Progress */}
          <div className="bg-white rounded-3xl shadow-md p-4">
            <div className="flex justify-between text-sm text-gray-500 mb-2">
              <span className="font-bold text-gray-700">Zeile {currentIndex + 1} / {totalItems}</span>
              <span>{isEmriNo && `İş Emri: ${isEmriNo}`}</span>
            </div>
            <div className="w-full bg-gray-100 rounded-full h-3">
              <div className="bg-blue-700 h-3 rounded-full transition-all"
                style={{ width: `${(currentIndex / totalItems) * 100}%` }} />
            </div>
          </div>

          {/* OCR */}
          <div className="bg-amber-50 border-2 border-amber-200 rounded-3xl p-5 space-y-3">
            <p className="text-xs font-bold text-amber-600 uppercase tracking-widest">📖 Zettel</p>
            <p className="text-2xl font-bold text-gray-900 leading-tight">{currentItem.urun_adi}</p>
            <p className="text-5xl font-black text-blue-900">{currentItem.miktar}</p>
          </div>

          {/* Match */}
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

              {stokSugg.length > 1 && (
                <div className="space-y-2">
                  <p className="text-sm text-gray-400">Andere Vorschläge:</p>
                  {stokSugg.slice(1, 4).map(s => (
                    <button key={s.stok_kodu} onClick={() => setEditStok(s)}
                      className="w-full text-left text-sm bg-gray-50 active:bg-blue-50 border border-gray-200 rounded-xl px-4 py-3">
                      <span className="font-mono text-xs text-gray-400">{s.stok_kodu}</span><br />{s.stok_adi}
                    </button>
                  ))}
                </div>
              )}

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
                {stokSugg.length > 0 && (
                  <div className="mt-2 space-y-1 max-h-52 overflow-y-auto">
                    {stokSugg.map(s => (
                      <button key={s.stok_kodu}
                        onClick={() => { setEditStok(s); setStokSearch(s.stok_adi); setStokSugg([]); }}
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

          {!isEditing && (
            <div className="grid grid-cols-2 gap-3">
              <button onClick={handleSkip}
                className="bg-gray-200 active:bg-gray-300 text-gray-800 font-bold py-6 rounded-3xl text-xl">
                ⏭ Überspringen
              </button>
              <button onClick={handleConfirm}
                className="bg-green-600 active:bg-green-800 text-white font-bold py-6 rounded-3xl text-2xl shadow-lg">
                ✅ Ja
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── MORE PAGES ───────────────────────────────────────────────────── */}
      {step === "more_pages" && (
        <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-6">
          <div className="text-6xl">📄</div>
          <h2 className="text-2xl font-bold text-gray-900">Weitere Seite?</h2>
          <p className="text-gray-500 text-lg">
            {allConfirmed.filter(i => !i.skipped).length} Produkte bestätigt
          </p>
          <div className="grid grid-cols-2 gap-4">
            <button onClick={() => { setScanResult(null); setPageConfirmed([]); setCurrentIndex(0); setPreviewImage(null); setStep("scan"); }}
              className="bg-green-600 active:bg-green-800 text-white font-bold py-8 rounded-3xl text-2xl shadow-lg">
              ✅ Ja
            </button>
            <button onClick={() => setStep("done")}
              className="bg-red-600 active:bg-red-800 text-white font-bold py-8 rounded-3xl text-2xl shadow-lg">
              ❌ Nein
            </button>
          </div>
          <p className="text-gray-400 text-sm">Nein = Excel wird erstellt</p>
        </div>
      )}

      {/* ── DONE ─────────────────────────────────────────────────────────── */}
      {step === "done" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-8 text-center space-y-4">
            <div className="text-7xl">✅</div>
            <h2 className="text-3xl font-bold text-green-700">Fertig!</h2>
            <p className="text-xl text-gray-600">
              <strong>{confirmedCount}</strong> Produkte — İş Emri {isEmriNo || "—"}
            </p>
          </div>

          <div className="bg-white rounded-3xl shadow-md p-5 space-y-2">
            <h3 className="font-bold text-gray-700 text-lg">Übersicht</h3>
            <div className="space-y-2 max-h-64 overflow-y-auto">
              {allConfirmed.map((item, idx) => (
                <div key={idx} className={`flex items-center gap-3 p-3 rounded-2xl ${item.skipped ? "bg-gray-50 text-gray-400" : "bg-green-50"}`}>
                  <span className="text-xl">{item.skipped ? "⏭" : "✅"}</span>
                  <div className="flex-1 min-w-0">
                    <p className="font-medium truncate text-base">{item.confirmed_stok?.stok_adi || item.original_urun_adi}</p>
                    {item.confirmed_stok && <p className="text-xs font-mono text-gray-400">{item.confirmed_stok.stok_kodu}</p>}
                  </div>
                  {!item.skipped && <span className="font-bold text-blue-900 text-xl whitespace-nowrap">{item.confirmed_miktar}</span>}
                </div>
              ))}
            </div>
          </div>

          <button onClick={handleExport}
            className="w-full bg-blue-700 active:bg-blue-900 text-white font-bold py-7 rounded-3xl text-2xl shadow-xl">
            📥 Excel Herunterladen
          </button>

          <button onClick={() => { setAllConfirmed([]); setScanResult(null); setPageConfirmed([]); setCurrentIndex(0); setPreviewImage(null); setIsEmriNo(""); setStep("scan"); }}
            className="w-full bg-gray-200 active:bg-gray-300 text-gray-700 font-bold py-5 rounded-3xl text-xl">
            🔄 Neu starten
          </button>
        </div>
      )}

      {/* ── UPDATE ───────────────────────────────────────────────────────── */}
      {step === "update" && (
        <div className="space-y-4">
          <div className="bg-white rounded-3xl shadow-md p-6 space-y-4">
            <h2 className="text-2xl font-bold text-blue-900">🔄 Stok aktualisieren</h2>
            <p className="text-gray-500">
              Neue Stok-Excel hochladen. Nur wirklich <strong>neue Produkte</strong> (neue STOK_KODU) werden hinzugefügt. Bestehende werden nicht überschrieben.
            </p>
            <div className="bg-blue-50 rounded-2xl p-4 text-sm text-blue-800">
              Aktuell: <strong>{stokData.length.toLocaleString()}</strong> Produkte im System
            </div>

            <label className="block cursor-pointer">
              <div className="border-2 border-dashed border-blue-400 rounded-2xl p-8 text-center active:bg-blue-50">
                <div className="text-5xl mb-3">📂</div>
                <p className="text-blue-700 font-bold text-xl">Excel hochladen</p>
                <p className="text-gray-400 mt-1">Neue Stok Kart Kayıtları.xlsx</p>
              </div>
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleUpdateExcel} />
            </label>

            {updateStatus && (
              <div className={`rounded-2xl p-4 text-base font-medium ${updateStatus.startsWith("✅") ? "bg-green-50 text-green-800" : updateStatus.startsWith("❌") ? "bg-red-50 text-red-800" : "bg-gray-50 text-gray-600"}`}>
                {updateStatus}
              </div>
            )}
          </div>

          <button onClick={() => setStep("scan")}
            className="w-full bg-gray-200 active:bg-gray-300 text-gray-800 font-bold py-5 rounded-3xl text-xl">
            ← Zurück
          </button>
        </div>
      )}
    </div>
  );
}
