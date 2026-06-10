import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import type { ExportPayload } from "@/lib/types";

// Generate Fiş No from current datetime: DDMMYYYYHHMMSS
function fisnо(): string {
  const n = new Date();
  return (
    String(n.getDate()).padStart(2, "0") +
    String(n.getMonth() + 1).padStart(2, "0") +
    String(n.getFullYear()) +
    String(n.getHours()).padStart(2, "0") +
    String(n.getMinutes()).padStart(2, "0") +
    String(n.getSeconds()).padStart(2, "0")
  );
}

export async function POST(req: NextRequest) {
  try {
    const payload: ExportPayload = await req.json();
    const { is_emri_no, tarih, items } = payload;

    const fisNo = fisnо();
    const confirmedItems = items.filter(i => !i.skipped && i.confirmed_stok);

    const header = [
      "Fiş No.(*)",
      "Tarih",
      "Belge Tipi",
      "İş Emri/Sip.",
      "Sipariş Kale",
      "Depo Öncel",
      "L. Depo(*)",
      "Çık. Depo(*)",
      "*Mamul Kod",
      "Fire Depo",
      "Miktar(*)",
      "2.Miktar",
      "Öncelik",
      "Açıklama",
      "Revizyon N",
      "Ek Alan-1",
      "Ek Alan-2",
      "Oto. Yarı M",
      "Oto. Yarı M (2)",
      "Bakiye (0/1)",
      "Mamüller Ölçü",
    ];

    const rows = confirmedItems.map(item => ({
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

    ws["!cols"] = [
      { wch: 16 }, // Fiş No
      { wch: 12 }, // Tarih
      { wch: 12 }, // Belge Tipi
      { wch: 14 }, // İş Emri
      { wch: 10 }, // Sipariş Kale
      { wch: 10 }, // Depo Öncel
      { wch: 10 }, // L. Depo
      { wch: 10 }, // Çık. Depo
      { wch: 18 }, // Mamul Kod
      { wch: 10 }, // Fire Depo
      { wch: 10 }, // Miktar
      ...Array(10).fill({ wch: 10 }),
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    const filename = `Uretim_${is_emri_no || fisNo}_${tarih.replace(/\./g, "-")}.xlsx`;

    return new NextResponse(buf, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : "Dışa aktarma başarısız";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
