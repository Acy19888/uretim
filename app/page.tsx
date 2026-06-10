"use client";

import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { findMatches } from "@/lib/fuzzy";
import type { StokItem, ScannedItem, ConfirmedItem, ScanResult } from "@/lib/types";

// ─── Step types ───────────────────────────────────────────────────────────────
type Step = "setup" | "scan" | "processing" | "confirm" | "done";

// ─── Helpers ──────────────────────────────────────────────────────────────────
function toBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.split(",")[1]);
    };
    reader.onerror = reject;
  });
}

function today(): string {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
}

// ─── Component ────────────────────────────────────────────────────────────────
export default function Home() {
  const [step, setStep] = useState<Step>("setup");
  const [stokData, setStokData] = useState<StokItem[]>([]);
  const [scanResult, setScanResult] = useState<ScanResult | null>(null);
  const [confirmedItems, setConfirmedItems] = useState<ConfirmedItem[]>([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [previewImage, setPreviewImage] = useState<string | null>(null);

  // Edit state for current item
  const [editMiktar, setEditMiktar] = useState("");
  const [editStok, setEditStok] = useState<StokItem | null>(null);
  const [stokSearch, setStokSearch] = useState("");
  const [stokSuggestions, setStokSuggestions] = useState<StokItem[]>([]);
  const [isEditing, setIsEditing] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const stokInputRef = useRef<HTMLInputElement>(null);

  // Load stok data from localStorage on mount
  useEffect(() => {
    const stored = localStorage.getItem("stok_data");
    if (stored) {
      try {
        const parsed: StokItem[] = JSON.parse(stored);
        if (parsed.length > 0) {
          setStokData(parsed);
          setStep("scan");
        }
      } catch {
        localStorage.removeItem("stok_data");
      }
    }
  }, []);

  // ── Setup: Parse uploaded Stok Kart Excel ──────────────────────────────────
  async function handleStokUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(null);

    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      // Try first sheet
      const sheetName = wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<Record<string, string>>(ws, {
        raw: false,
        defval: "",
      });

      const items: StokItem[] = [];
      for (const row of rows) {
        // Handle different possible column names
        const kodu =
          row["STOK_KODU"] ||
          row["Stok Kodu"] ||
          row["STOK KODU"] ||
          row["stok_kodu"] ||
          "";
        const adi =
          row["STOK_ADI"] ||
          row["Stok Adı"] ||
          row["STOK ADI"] ||
          row["stok_adi"] ||
          "";
        const uretim =
          row["Üretim"] ||
          row["URETIM"] ||
          row["üretim"] ||
          "";
        const cesit =
          row["Çeşit"] ||
          row["CESIT"] ||
          row["çeşit"] ||
          "";

        if (kodu && adi) {
          items.push({
            stok_kodu: String(kodu).trim(),
            stok_adi: String(adi).trim(),
            uretim: String(uretim).trim(),
            cesit: String(cesit).trim(),
          });
        }
      }

      if (items.length === 0) {
        setError(
          "Keine Produkte gefunden. Bitte prüfe ob die Excel die Spalten STOK_KODU und STOK_ADI hat."
        );
        return;
      }

      localStorage.setItem("stok_data", JSON.stringify(items));
      setStokData(items);
      setStep("scan");
    } catch (err) {
      setError("Fehler beim Lesen der Excel-Datei: " + String(err));
    }
  }

  // ── Scan: Upload/take photo ────────────────────────────────────────────────
  async function handleImageSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError(null);

    // Show preview
    const url = URL.createObjectURL(file);
    setPreviewImage(url);
    setStep("processing");

    try {
      const base64 = await toBase64(file);
      const res = await fetch("/api/scan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ imageBase64: base64, mimeType: file.type }),
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error || "Scan fehlgeschlagen");
      }

      const result: ScanResult = await res.json();

      if (!result.items || result.items.length === 0) {
        throw new Error("Keine Produkte erkannt. Bitte Foto erneut aufnehmen.");
      }

      setScanResult(result);
      setCurrentIndex(0);
      setConfirmedItems([]);
      prepareCurrentItem(result.items[0]);
      setStep("confirm");
    } catch (err) {
      setError(String(err));
      setStep("scan");
    }
  }

  // ── Confirm: Prepare each item ─────────────────────────────────────────────
  const prepareCurrentItem = useCallback(
    (item: ScannedItem) => {
      setEditMiktar(item.miktar || "");
      setIsEditing(false);
      setStokSearch(item.urun_adi || "");

      if (stokData.length > 0) {
        const matches = findMatches(item.urun_adi, stokData, 5);
        setStokSuggestions(matches);
        setEditStok(matches[0] || null);
      }
    },
    [stokData]
  );

  useEffect(() => {
    if (
      step === "confirm" &&
      scanResult &&
      currentIndex < scanResult.items.length
    ) {
      prepareCurrentItem(scanResult.items[currentIndex]);
    }
  }, [currentIndex, step, scanResult, prepareCurrentItem]);

  // ── Confirm: User clicks Ja ────────────────────────────────────────────────
  function handleConfirm() {
    const currentItem = scanResult!.items[currentIndex];
    const qty = parseFloat(editMiktar.replace(",", ".")) || 0;

    const confirmed: ConfirmedItem = {
      original_urun_adi: currentItem.urun_adi,
      original_miktar: currentItem.miktar,
      confirmed_stok: editStok,
      confirmed_miktar: qty,
      skipped: false,
    };

    const updated = [...confirmedItems, confirmed];
    setConfirmedItems(updated);
    advanceOrFinish(updated);
  }

  // ── Confirm: User clicks Überspringen ─────────────────────────────────────
  function handleSkip() {
    const currentItem = scanResult!.items[currentIndex];
    const skipped: ConfirmedItem = {
      original_urun_adi: currentItem.urun_adi,
      original_miktar: currentItem.miktar,
      confirmed_stok: null,
      confirmed_miktar: 0,
      skipped: true,
    };
    const updated = [...confirmedItems, skipped];
    setConfirmedItems(updated);
    advanceOrFinish(updated);
  }

  function advanceOrFinish(items: ConfirmedItem[]) {
    const nextIdx = currentIndex + 1;
    if (nextIdx >= scanResult!.items.length) {
      setConfirmedItems(items);
      setStep("done");
    } else {
      setCurrentIndex(nextIdx);
    }
  }

  // ── Stok search in edit mode ───────────────────────────────────────────────
  function handleStokSearchChange(q: string) {
    setStokSearch(q);
    if (q.length >= 2) {
      const matches = findMatches(q, stokData, 8);
      setStokSuggestions(matches);
    } else {
      setStokSuggestions([]);
    }
  }

  // ── Export Excel ───────────────────────────────────────────────────────────
  async function handleExport() {
    if (!scanResult) return;
    setError(null);

    try {
      const res = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          is_emri_no: scanResult.is_emri_no,
          tarih: scanResult.tarih || today(),
          items: confirmedItems,
        }),
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.error);
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `UretimCikti_${scanResult.is_emri_no}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      setError("Export fehlgeschlagen: " + String(err));
    }
  }

  const currentItem = scanResult?.items[currentIndex];
  const totalItems = scanResult?.items.length || 0;
  const confirmedCount = confirmedItems.filter((i) => !i.skipped).length;

  // ─────────────────────────────────────────────────────────────────────────
  // RENDER
  // ─────────────────────────────────────────────────────────────────────────

  return (
    <div className="space-y-4">
      {/* Error banner */}
      {error && (
        <div className="bg-red-50 border border-red-300 text-red-800 rounded-xl p-4 flex items-start gap-3">
          <span className="text-xl">⚠️</span>
          <div>
            <p className="font-bold">Fehler</p>
            <p className="text-sm">{error}</p>
          </div>
          <button
            className="ml-auto text-red-500 hover:text-red-800"
            onClick={() => setError(null)}
          >
            ✕
          </button>
        </div>
      )}

      {/* ── STEP: SETUP ────────────────────────────────────────────────────── */}
      {step === "setup" && (
        <div className="card space-y-5">
          <h2 className="text-2xl font-bold text-blue-900">
            🔧 Stok Verisi Yükle
          </h2>
          <p className="text-gray-600">
            İlk kullanımda <strong>Stok Kart Kayıtları.xlsx</strong> dosyasını
            yükleyiniz. Veriler tarayıcıda saklanır.
          </p>
          <label className="block cursor-pointer">
            <div className="border-2 border-dashed border-blue-400 rounded-xl p-8 text-center hover:bg-blue-50 transition-colors">
              <div className="text-5xl mb-3">📂</div>
              <p className="text-blue-700 font-bold text-lg">
                Excel Dosyası Seç
              </p>
              <p className="text-gray-500 text-sm mt-1">
                Stok Kart Kayıtları.xlsx
              </p>
            </div>
            <input
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={handleStokUpload}
            />
          </label>
          {stokData.length > 0 && (
            <p className="text-green-600 font-medium text-center">
              ✅ {stokData.length} Produkt geladen
            </p>
          )}
        </div>
      )}

      {/* ── STEP: SCAN ─────────────────────────────────────────────────────── */}
      {step === "scan" && (
        <div className="space-y-4">
          <div className="card space-y-2">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold text-blue-900">
                📷 Zettel Scannen
              </h2>
              <span className="text-sm text-green-600 font-medium bg-green-50 px-3 py-1 rounded-full">
                ✅ {stokData.length} Produkt geladen
              </span>
            </div>
            <p className="text-gray-500 text-sm">
              Produktionszettel fotografieren oder Bild auswählen
            </p>
          </div>

          {/* Camera button */}
          <label className="block cursor-pointer">
            <div className="bg-blue-700 hover:bg-blue-800 active:scale-95 text-white rounded-2xl p-8 text-center transition-all shadow-lg">
              <div className="text-6xl mb-4">📷</div>
              <p className="text-2xl font-bold">Kamera</p>
              <p className="text-blue-200 mt-1">Foto aufnehmen</p>
            </div>
            <input
              type="file"
              accept="image/*"
              capture="environment"
              className="hidden"
              onChange={handleImageSelect}
            />
          </label>

          {/* File upload */}
          <label className="block cursor-pointer">
            <div className="bg-white hover:bg-gray-50 active:scale-95 border-2 border-gray-200 rounded-2xl p-6 text-center transition-all">
              <div className="text-4xl mb-3">🖼️</div>
              <p className="text-lg font-bold text-gray-700">Bild auswählen</p>
              <p className="text-gray-400 text-sm">aus Galerie</p>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept="image/*"
              className="hidden"
              onChange={handleImageSelect}
            />
          </label>

          {/* Change stok data */}
          <button
            className="w-full text-center text-gray-400 text-sm underline py-2"
            onClick={() => setStep("setup")}
          >
            Stok Kart Datei wechseln
          </button>
        </div>
      )}

      {/* ── STEP: PROCESSING ───────────────────────────────────────────────── */}
      {step === "processing" && (
        <div className="card text-center space-y-6 py-12">
          {previewImage && (
            // eslint-disable-next-line @next/next/no-img-element
            <img
              src={previewImage}
              alt="Zettel Vorschau"
              className="max-h-48 mx-auto rounded-xl shadow object-contain"
            />
          )}
          <div className="animate-spin text-5xl">⚙️</div>
          <div>
            <p className="text-xl font-bold text-blue-900">
              Zettel wird analysiert...
            </p>
            <p className="text-gray-500 text-sm mt-1">
              KI liest die Handschrift
            </p>
          </div>
        </div>
      )}

      {/* ── STEP: CONFIRM ──────────────────────────────────────────────────── */}
      {step === "confirm" && currentItem && (
        <div className="space-y-4">
          {/* Progress */}
          <div className="card">
            <div className="flex justify-between text-sm text-gray-500 mb-2">
              <span>
                Zeile {currentIndex + 1} / {totalItems}
              </span>
              <span>
                {confirmedCount} bestätigt, {confirmedItems.filter((i) => i.skipped).length} übersprungen
              </span>
            </div>
            <div className="w-full bg-gray-200 rounded-full h-2.5">
              <div
                className="bg-blue-700 h-2.5 rounded-full transition-all"
                style={{
                  width: `${((currentIndex) / totalItems) * 100}%`,
                }}
              />
            </div>
          </div>

          {/* Header info */}
          {scanResult && (
            <div className="flex gap-2 text-sm">
              <span className="bg-blue-100 text-blue-800 px-3 py-1 rounded-full font-medium">
                İş Emri: {scanResult.is_emri_no || "—"}
              </span>
              <span className="bg-gray-100 text-gray-700 px-3 py-1 rounded-full">
                {scanResult.tarih || today()}
              </span>
            </div>
          )}

          {/* OCR result */}
          <div className="card space-y-4">
            <div className="text-xs font-bold text-gray-400 uppercase tracking-wider">
              📖 Zettelinhalt (KI-Erkennung)
            </div>
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-4">
              <p className="text-gray-500 text-xs mb-1">Ürün Adı</p>
              <p className="text-xl font-bold text-gray-900">
                {currentItem.urun_adi}
              </p>
              <p className="text-gray-500 text-xs mt-3 mb-1">Miktar</p>
              <p className="text-3xl font-bold text-blue-900">
                {currentItem.miktar}
              </p>
            </div>
          </div>

          {/* Matched product */}
          <div className="card space-y-4">
            <div className="text-xs font-bold text-gray-400 uppercase tracking-wider">
              🔍 Stok Eşleşme
            </div>

            {!isEditing ? (
              <>
                {editStok ? (
                  <div className="bg-green-50 border border-green-200 rounded-xl p-4">
                    <p className="text-xs text-gray-500 mb-1">Stok Kodu</p>
                    <p className="font-mono font-bold text-green-800 text-lg">
                      {editStok.stok_kodu}
                    </p>
                    <p className="text-xs text-gray-500 mt-2 mb-1">Ürün Adı</p>
                    <p className="font-bold text-gray-900">{editStok.stok_adi}</p>
                    {editStok.cesit && (
                      <span className="mt-2 inline-block text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full">
                        {editStok.cesit}
                      </span>
                    )}
                  </div>
                ) : (
                  <div className="bg-red-50 border border-red-200 rounded-xl p-4 text-center">
                    <p className="text-red-600 font-medium">
                      Kein passendes Produkt gefunden
                    </p>
                  </div>
                )}

                {/* Other suggestions */}
                {stokSuggestions.length > 1 && (
                  <div className="space-y-2">
                    <p className="text-xs text-gray-500">Andere Vorschläge:</p>
                    {stokSuggestions.slice(1, 4).map((s) => (
                      <button
                        key={s.stok_kodu}
                        onClick={() => setEditStok(s)}
                        className="w-full text-left text-sm bg-gray-50 hover:bg-blue-50 border border-gray-200 rounded-lg px-3 py-2 transition-colors"
                      >
                        <span className="font-mono text-xs text-gray-400">
                          {s.stok_kodu}
                        </span>{" "}
                        — {s.stok_adi}
                      </button>
                    ))}
                  </div>
                )}

                {/* Quantity display */}
                <div className="flex items-center gap-3">
                  <div className="flex-1">
                    <p className="text-xs text-gray-500 mb-1">
                      Bestätigte Menge
                    </p>
                    <p className="text-4xl font-bold text-blue-900">
                      {editMiktar}
                    </p>
                  </div>
                  <button
                    onClick={() => setIsEditing(true)}
                    className="text-sm text-blue-600 underline"
                  >
                    Bearbeiten
                  </button>
                </div>
              </>
            ) : (
              /* Edit mode */
              <div className="space-y-4">
                <div>
                  <label className="text-sm font-bold text-gray-600 block mb-2">
                    Menge korrigieren
                  </label>
                  <input
                    type="number"
                    value={editMiktar}
                    onChange={(e) => setEditMiktar(e.target.value)}
                    className="w-full border-2 border-blue-300 rounded-xl p-3 text-2xl font-bold text-center focus:outline-none focus:border-blue-600"
                    autoFocus
                  />
                </div>

                <div>
                  <label className="text-sm font-bold text-gray-600 block mb-2">
                    Produkt suchen
                  </label>
                  <input
                    ref={stokInputRef}
                    type="text"
                    value={stokSearch}
                    onChange={(e) => handleStokSearchChange(e.target.value)}
                    placeholder="z.B. ATLAS KOCAK oder Stok Kodu..."
                    className="w-full border-2 border-gray-300 rounded-xl p-3 focus:outline-none focus:border-blue-500"
                  />
                  {stokSuggestions.length > 0 && (
                    <div className="mt-2 space-y-1 max-h-48 overflow-y-auto">
                      {stokSuggestions.map((s) => (
                        <button
                          key={s.stok_kodu}
                          onClick={() => {
                            setEditStok(s);
                            setStokSearch(s.stok_adi);
                            setStokSuggestions([]);
                          }}
                          className={`w-full text-left text-sm rounded-lg px-3 py-2 transition-colors border ${
                            editStok?.stok_kodu === s.stok_kodu
                              ? "bg-green-50 border-green-300 font-bold"
                              : "bg-gray-50 border-gray-200 hover:bg-blue-50"
                          }`}
                        >
                          <span className="font-mono text-xs text-gray-400">
                            {s.stok_kodu}
                          </span>
                          <br />
                          {s.stok_adi}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <button
                  onClick={() => setIsEditing(false)}
                  className="btn-primary w-full"
                >
                  Speichern
                </button>
              </div>
            )}
          </div>

          {/* Action buttons */}
          {!isEditing && (
            <div className="grid grid-cols-2 gap-3">
              <button onClick={handleSkip} className="btn-secondary">
                ⏭ Überspringen
              </button>
              <button onClick={handleConfirm} className="btn-success">
                ✅ Bestätigen
              </button>
            </div>
          )}
        </div>
      )}

      {/* ── STEP: DONE ─────────────────────────────────────────────────────── */}
      {step === "done" && (
        <div className="space-y-4">
          <div className="card text-center space-y-4">
            <div className="text-6xl">✅</div>
            <h2 className="text-2xl font-bold text-green-700">
              Fertig!
            </h2>
            <p className="text-gray-600">
              <strong>{confirmedCount}</strong> Produkte bestätigt,{" "}
              <strong>{confirmedItems.filter((i) => i.skipped).length}</strong>{" "}
              übersprungen
            </p>
            <div className="text-sm text-gray-500 bg-gray-50 rounded-xl p-3">
              <p>İş Emri: <strong>{scanResult?.is_emri_no}</strong></p>
              <p>Tarih: <strong>{scanResult?.tarih || today()}</strong></p>
            </div>
          </div>

          {/* Summary of confirmed items */}
          <div className="card space-y-3">
            <h3 className="font-bold text-gray-700">Bestätigte Produkte</h3>
            <div className="space-y-2 max-h-64 overflow-y-auto">
              {confirmedItems.map((item, idx) => (
                <div
                  key={idx}
                  className={`flex items-center gap-3 p-3 rounded-lg text-sm ${
                    item.skipped
                      ? "bg-gray-50 text-gray-400 line-through"
                      : "bg-green-50"
                  }`}
                >
                  <span className="text-lg">
                    {item.skipped ? "⏭" : "✅"}
                  </span>
                  <div className="flex-1 min-w-0">
                    <p className="font-medium truncate">
                      {item.confirmed_stok?.stok_adi || item.original_urun_adi}
                    </p>
                    {item.confirmed_stok && (
                      <p className="text-xs font-mono text-gray-500">
                        {item.confirmed_stok.stok_kodu}
                      </p>
                    )}
                  </div>
                  {!item.skipped && (
                    <span className="font-bold text-blue-900 text-lg whitespace-nowrap">
                      {item.confirmed_miktar} St.
                    </span>
                  )}
                </div>
              ))}
            </div>
          </div>

          <button onClick={handleExport} className="btn-primary w-full text-xl py-5">
            📥 Excel Herunterladen
          </button>

          <button
            onClick={() => {
              setScanResult(null);
              setConfirmedItems([]);
              setCurrentIndex(0);
              setPreviewImage(null);
              setStep("scan");
            }}
            className="btn-secondary w-full"
          >
            🔄 Neuen Zettel scannen
          </button>
        </div>
      )}
    </div>
  );
}
