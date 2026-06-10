import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

export async function POST(req: NextRequest) {
  try {
    const { imageBase64, mimeType } = await req.json();

    if (!imageBase64) {
      return NextResponse.json({ error: "Kein Bild erhalten" }, { status: 400 });
    }

    const prompt = `Du analysierst einen handgeschriebenen Produktionszettel einer Fenster/Tür-Fabrik (WINDOFORM).

Der Zettel ist ein "Yarı Mamul Giriş ve Çıkış Defteri" (Halbfertigware Ein-/Ausgangsbuch).

Extrahiere alle Informationen und gib sie als JSON zurück.

**Wichtige Spalten:**
- İş Emri No: Die Auftragsnummer (oben auf dem Zettel)
- Sıra No: Die Laufende Nummer
- Tarih: Das Datum
- Ürün Adı: Der Produktname (linke Spalte in der Tabelle)
- Miktar: Die Stückzahl/Menge (rechte Spalte(n) in der Tabelle)

**Regeln:**
- Lies alle beschriebenen Zeilen
- Falls eine Zeile mehrere Zahlen hat, nimm die letzte/größte als Gesamtmenge
- Leere Zeilen überspringen
- Handschrift kann schwer lesbar sein - gib dein Bestes
- Zahlen können auch als Kombinationen stehen wie "1192" oder "280"

Antworte NUR mit diesem JSON (kein Markdown, kein Text davor/danach):
{
  "is_emri_no": "...",
  "tarih": "...",
  "sira_no": "...",
  "items": [
    {"urun_adi": "...", "miktar": "..."},
    ...
  ]
}`;

    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "user",
          content: [
            {
              type: "image_url",
              image_url: {
                url: `data:${mimeType || "image/jpeg"};base64,${imageBase64}`,
                detail: "high",
              },
            },
            {
              type: "text",
              text: prompt,
            },
          ],
        },
      ],
      max_tokens: 2000,
      temperature: 0.1,
    });

    const content = response.choices[0]?.message?.content || "";

    // Clean up the response - remove markdown code blocks if present
    const cleaned = content
      .replace(/```json\n?/g, "")
      .replace(/```\n?/g, "")
      .trim();

    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      console.error("JSON parse error:", cleaned);
      return NextResponse.json(
        { error: "KI konnte den Zettel nicht lesen. Bitte erneut versuchen." },
        { status: 422 }
      );
    }

    return NextResponse.json(parsed);
  } catch (error: unknown) {
    console.error("Scan error:", error);
    const message =
      error instanceof Error ? error.message : "Unbekannter Fehler";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
