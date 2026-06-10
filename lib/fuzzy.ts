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
function scoreMatch(query: string, stokAdi: string): number {
  const normQuery = normalize(query);
  const normStok = normalize(stokAdi);

  const queryTokens = normQuery.split(" ").filter((t) => t.length > 1);
  const stokTokens = normStok.split(" ").filter((t) => t.length > 1);

  if (queryTokens.length === 0) return 0;

  let score = 0;

  for (const qt of queryTokens) {
    for (const st of stokTokens) {
      if (st === qt) {
        score += qt.length * 3; // exact token match = high score
      } else if (st.startsWith(qt) || qt.startsWith(st)) {
        score += Math.min(qt.length, st.length) * 2;
      } else if (st.includes(qt) || qt.includes(st)) {
        score += Math.min(qt.length, st.length);
      }
    }
  }

  // Bonus if the normalized product name contains the whole query
  if (normStok.includes(normQuery)) {
    score += normQuery.length * 2;
  }

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
      score: scoreMatch(query, item.stok_adi),
    }))
    .filter((s) => s.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, topN);

  return scored.map((s) => s.item);
}
