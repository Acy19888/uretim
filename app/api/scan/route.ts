import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export async function POST(req: NextRequest) {
  try {
    const openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    const { imageBase64, mimeType } = await req.json();

    if (!imageBase64) {
      return NextResponse.json({ error: "Görsel alınamadı" }, { status: 400 });
    }

    const prompt = `Bir pencere/kapı fabrikasına (WINDOFORM) ait el yazısıyla doldurulmuş üretim fişini analiz ediyorsun.

Fiş, bir "Yarı Mamul Giriş ve Çıkış Defteri" sayfasıdır.

Tüm bilgileri çıkar ve JSON formatında döndür.

**Önemli Sütunlar:**
- İş Emri No: Sipariş numarası (fişin üst kısmında)
- Sıra No: Sıra numarası
- Tarih: Tarih
- RAL Renk: Sayfanın en üstünde büyük yazılı 4 haneli RAL renk kodu (örn: 9005, 8019, 1013). Bu çok önemli — her sayfada bir tane olur.
- Ürün Adı: Ürün adı (tablonun sol sütunu)
- Miktar: Adet/miktar (tablonun sağ sütunu)

**Kurallar:**
- Tüm doldurulmuş satırları oku
- Boş satırları atla
- El yazısı okunması zor olabilir - elinden gelenin en iyisini yap
- Miktar sütununda "+" işareti varsa (örn: "1848 + 1900") tüm sayıları topla ve tek bir sayı olarak yaz (örn: "3748")
- Miktar sütununda birden fazla sayı ama "+" yoksa en büyük sayıyı al

YALNIZCA şu JSON formatında yanıt ver (önünde/arkasında Markdown veya metin olmasın):
{
  "is_emri_no": "...",
  "tarih": "...",
  "sira_no": "...",
  "ral_renk": "9005",
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
        { error: "Yapay zeka fişi okuyamadı. Lütfen tekrar deneyin." },
        { status: 422 }
      );
    }

    return NextResponse.json(parsed);
  } catch (error: unknown) {
    console.error("Scan error:", error);
    const message =
      error instanceof Error ? error.message : "Bilinmeyen hata";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
