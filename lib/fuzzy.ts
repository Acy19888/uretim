import type { StokItem } from "./types";

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

  let score = 0;
  for (const qt of queryTokens) {
    for (const st of adiTokens) {
      if (st === qt)                          score += qt.length * 3;
      else if (st.startsWith(qt) || qt.startsWith(st)) score += Math.min(qt.length, st.length) * 2;
      else if (st.includes(qt) || qt.includes(st))     score += Math.min(qt.length, st.length);
    }
  }

  // Bonus: tüm query ürün adında geçiyorsa
  if (normAdi.includes(normQuery)) score += normQuery.length * 2;

  return score;
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
