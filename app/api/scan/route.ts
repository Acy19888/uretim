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

**Miktar Sütunları Hakkında ÇOK ÖNEMLİ:**
- Bu fişte "MİKTAR", "MİKTAR", "MİKTAR", "TOP. MİK." şeklinde birden fazla miktar sütunu var
- Bir satırda birden fazla miktar sütununa değer yazılmış olabilir (örn: birinci sütunda "560", ikinci sütunda "468")
- Aynı satırın birden fazla miktar sütununda değer varsa veya "+" işareti varsa → HEPSİNİ TOPLA
- Örnekler:
  - "1848" ve "1900" aynı satırda (farklı sütunlarda veya "+" ile) → miktar: "3748"
  - "560" ve "468" aynı satırda → miktar: "1028"
  - Sadece tek sütunda "5712" → miktar: "5712"
- "TOP. MİK." sütununda zaten toplam yazılıysa onu kullan
- Aksi hâlde tüm miktar sütunlarını kendin topla

**Kısaltmalar — Ürün Adında Geçen Harfler:**
- Ürün adının sonunda tek başına **E** → **Erkek** (örn: "WC geniş Ayna E" → "WC geniş Ayna Erkek")
- Ürün adının sonunda tek başına **D** → **Dişi** (örn: "WC geniş Ayna D" → "WC geniş Ayna Disi")
- **HBSB** veya benzer yazım (HB5B vb.) → **Hebe Schiebe**

**Tekrarlama İşareti (Ditto / " " ") ÇOK ÖNEMLİ:**
- Bazen bir satırda ürün adı yerine tırnak işareti veya ditto işareti yazılır: ", '', ,, veya birden fazla " " "
- Bu işaret "yukarıdaki satırın aynısı" anlamına gelir
- O satırın ürün adını bir üstteki satırla aynı yap — ama miktar ve varsa ek kelime (sol/sağ, iç/dış vb.) farklı olabilir
- Örnek:
  - Satır 1: "Atlas kıvrık kol sol" → 800
  - Satır 2: " " " sağ → 600
  - Sonuç: Satır 2 = "Atlas kıvrık kol sağ", miktar 600
- Ditto satırında ek bir kelime varsa (sol/sağ/iç/dış/büyük/küçük vb.) → onu ürün adına ekle
- Ditto işareti sadece ürün adı için geçerlidir; miktar her zaman o satırda yazılı değerdir

**Diğer Kurallar:**
- Tüm doldurulmuş satırları oku
- Boş satırları atla
- Üzeri çizili satırları atla (iptal edilmiş)
- El yazısı okunması zor olabilir - elinden gelenin en iyisini yap

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
