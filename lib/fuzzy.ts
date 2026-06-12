import type { StokItem } from "./types";

/**
 * Bigram benzerliği — yazım hatalarını tolere eder.
 * "BERLGRAT" ↔ "BELGIRAT" → ~0.57 (eşik: 0.5)
 */
function bigramSim(a: string, b: string): number {
  if (a.length < 2 || b.length < 2) return a === b ? 1 : 0;
  const bigramsA = new Set<string>();
  for (let i = 0; i < a.length - 1; i++) bigramsA.add(a.slice(i, i + 2));
  let hits = 0;
  for (let i = 0; i < b.length - 1; i++) {
    if (bigramsA.has(b.slice(i, i + 2))) hits++;
  }
  return (2 * hits) / (a.length + b.length - 2);
}

/**
 * Bilinen kısaltmaları açık hale getirir.
 * "Atlas HBSB" → "Atlas Hebe Schiebe"
 */
export const KISALTMALAR: Record<string, string> = {
  "HBSB": "Hebe Schiebe",
  "HST":  "Hebe Schiebe",
  "PVC":  "PVC",
  "AL":   "Aluminyum",
};

export function kisaltmaAc(text: string): string {
  let result = text;
  // Exact abbreviation map
  for (const [kisalt, acik] of Object.entries(KISALTMALAR)) {
    const re = new RegExp(`\\b${kisalt}\\b`, "gi");
    result = result.replace(re, acik);
  }
  // Pattern: HB?B — handles typos like HB5B, HBXB, HBAB etc. → Hebe Schiebe
  result = result.replace(/\bHB.B\b/gi, "Hebe Schiebe");
  // E / D am Ende → Erkek / Disi (nur als letztes Wort, sonst zu viele Fehlmatches)
  result = result.replace(/\bE$/i, "Erkek");
  result = result.replace(/\bD$/i, "Disi");
  return result;
}

// Normalize Turkish characters for comparison
export function normalize(str: string): string {
  return str
    .toUpperCase()
    .replace(/İ/g, "I")
    .replace(/Ğ/g, "G")
    .replace(/Ş/g, "S")
    .replace(/Ü/g, "U")
    .replace(/Ö/g, "O")
    .replace(/Ç/g, "C")
    .replace(/[^A-Z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Score how well a query matches a product name
function scoreMatch(query: string, stokAdi: string, stokKodu: string): number {
  const normQuery = normalize(query);
  const normAdi   = normalize(stokAdi);
  const normKodu  = stokKodu.toUpperCase().trim();

  // ── Stok kodu direkt eşleşme (kullanıcı kod yazıyorsa) ──
  if (normKodu === normQuery)         return 10000;
  if (normKodu.startsWith(normQuery)) return 8000 + normQuery.length * 10;
  if (normKodu.includes(normQuery))   return 6000 + normQuery.length * 5;

  // ── Token-tabanlı isim eşleşmesi ──
  const queryTokens = normQuery.split(" ").filter((t) => t.length > 1);
  const adiTokens   = normAdi.split(" ").filter((t) => t.length > 1);

  if (queryTokens.length === 0) return 0;

  let matchedScore = 0;
  let matchedTokenCount = 0;

  for (const qt of queryTokens) {
    let bestForToken = 0;
    for (const st of adiTokens) {
      let s = 0;
      if (st === qt)                                    s = qt.length * 3;
      else if (st.startsWith(qt) || qt.startsWith(st)) s = Math.min(qt.length, st.length) * 2;
      else if (st.includes(qt) || qt.includes(st))     s = Math.min(qt.length, st.length);
      else if (qt.length >= 4 && st.length >= 4) {
        // Yazım hatası toleransı: bigram benzerliği ≥ 0.5
        const sim = bigramSim(qt, st);
        if (sim >= 0.5) s = Math.min(qt.length, st.length) * sim * 1.5;
      }
      if (s > bestForToken) bestForToken = s;
    }
    if (bestForToken > 0) matchedTokenCount++;
    matchedScore += bestForToken;
  }

  // Bonus: tüm query ürün adında birebir geçiyorsa
  if (normAdi.includes(normQuery)) matchedScore += normQuery.length * 2;

  // Kesinlik bonusu: sorgu tokenları ürün tokenlarının büyük bölümünü kapsıyorsa
  // (örn: "EGE AKUSTIK" sorgusu için 6 tokenli JUMBO ürün yerine 5 tokenli sade ürün öne çıkar)
  const precision = matchedTokenCount / adiTokens.length;
  matchedScore = matchedScore * (1 + precision * 0.5);

  return matchedScore;
}

// Return top N matches for a query against the stok list
export function findMatches(
  query: string,
  items: StokItem[],
  topN = 5
): StokItem[] {
  if (!query || items.length === 0) return [];

  const scored = items
    .map((item) => ({
      item,
      score: scoreMatch(query, item.stok_adi, item.stok_kodu),
    }))
    .filter((s) => s.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, topN);

  return scored.map((s) => s.item);
}
