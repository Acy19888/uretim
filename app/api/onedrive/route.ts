import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import type { ExportPayload } from "@/lib/types";

const TENANT_ID     = process.env.AZURE_TENANT_ID     || "";
const CLIENT_ID     = process.env.AZURE_CLIENT_ID     || "";
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || "";
const TARGET_USER   = process.env.ONEDRIVE_USER_EMAIL || ""; // kimin OneDrive'ına
const FOLDER        = "WINDOFORM/Uretim";

// ── Access token al (Client Credentials — login gerektirmez) ──────────────
async function getToken(): Promise<string> {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
      }),
    }
  );
  if (!res.ok) throw new Error("Token alınamadı: " + (await res.text()));
  return (await res.json()).access_token;
}

// ── Tek dosyayı OneDrive'a yükle ─────────────────────────────────────────
async function uploadFile(
  token: string,
  filename: string,
  body: Buffer | Uint8Array,
  contentType: string
): Promise<void> {
  const path = `${FOLDER}/${filename}`;
  const url  = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(TARGET_USER)}/drive/root:/${encodeURIComponent(path)}:/content`;

  const res = await fetch(url, {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": contentType },
    body,
  });

  if (!res.ok) {
    let msg = `OneDrive yükleme hatası (${res.status})`;
    try { msg = (await res.json()).error?.message || msg; } catch { /* ignore */ }
    throw new Error(msg);
  }
}

// ── Excel oluştur (export route ile aynı mantık) ─────────────────────────
function fisnо(): string {
  const n = new Date();
  return (
    String(n.getDate()).padStart(2,"0") +
    String(n.getMonth()+1).padStart(2,"0") +
    String(n.getFullYear()) +
    String(n.getHours()).padStart(2,"0") +
    String(n.getMinutes()).padStart(2,"0") +
    String(n.getSeconds()).padStart(2,"0")
  );
}

function buildExcel(payload: ExportPayload): Buffer {
  const { tarih, items } = payload;
  const fisNo = fisnо();
  const confirmed = items.filter(i => !i.skipped && i.confirmed_stok);

  const header = [
    "Fiş No.(*)","Tarih","Belge Tipi","İş Emri/Sip.","Sipariş Kale",
    "Depo Öncel","L. Depo(*)","Çık. Depo(*)","*Mamul Kod","Fire Depo",
    "Miktar(*)","2.Miktar","Öncelik","Açıklama","Revizyon N",
    "Ek Alan-1","Ek Alan-2","Oto. Yarı M","Oto. Yarı M (2)","Bakiye (0/1)","Mamüller Ölçü",
  ];

  const rows = confirmed.map(item => ({
    "Fiş No.(*)":      fisNo,
    "Tarih":           tarih,
    "Belge Tipi":      "",
    "İş Emri/Sip.":   "",
    "Sipariş Kale":    "",
    "Depo Öncel":      "",
    "L. Depo(*)":      "500",
    "Çık. Depo(*)":    "500",
    "*Mamul Kod":      item.confirmed_stok?.stok_kodu || "",
    "Fire Depo":       "",
    "Miktar(*)":       item.confirmed_miktar,
    "2.Miktar":        "",
    "Öncelik":         "",
    "Açıklama":        "",
    "Revizyon N":      "",
    "Ek Alan-1":       "",
    "Ek Alan-2":       "",
    "Oto. Yarı M":     "",
    "Oto. Yarı M (2)": "",
    "Bakiye (0/1)":    "",
    "Mamüller Ölçü":   "",
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows, { header });
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

// ── POST handler ──────────────────────────────────────────────────────────
export async function POST(req: NextRequest) {
  try {
    if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !TARGET_USER) {
      return NextResponse.json(
        { error: "OneDrive env değişkenleri eksik (AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, ONEDRIVE_USER_EMAIL)" },
        { status: 500 }
      );
    }

    const { exportPayload, photos, baseName } = await req.json() as {
      exportPayload: ExportPayload;
      photos: { data: string; mime: string }[];
      baseName: string;
    };

    const token = await getToken();
    const uploads: Promise<void>[] = [];

    // Excel
    const excelBuf  = buildExcel(exportPayload);
    const excelName = `Uretim_${baseName}.xlsx`;
    uploads.push(uploadFile(token, excelName, excelBuf,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

    // Fotoğraflar
    for (let i = 0; i < photos.length; i++) {
      const { data, mime } = photos[i];
      const ext  = mime.split("/")[1]?.split(";")[0] || "jpg";
      const name = `Uretim_${baseName}_foto${photos.length > 1 ? `_${i+1}` : ""}.${ext}`;
      const bin  = Buffer.from(data, "base64");
      uploads.push(uploadFile(token, name, bin, mime));
    }

    await Promise.all(uploads);

    return NextResponse.json({
      ok: true,
      uploaded: uploads.length,
      folder: `OneDrive/${FOLDER}`,
    });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : "Bilinmeyen hata";
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
