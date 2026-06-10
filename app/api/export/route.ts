import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import type { ExportPayload } from "@/lib/types";

export async function POST(req: NextRequest) {
  try {
    const payload: ExportPayload = await req.json();

    const { is_emri_no, tarih, items } = payload;

    // Build rows matching the "boş Üsk" format
    // Columns: Fiş No(*), Tarih, Belge Tipi, İş Emri/Sip., Sipariş Kale, Depo Öncel,
    //          L. Depo(*), Çık. Depo(*), *Mamul Kod, Fire Depo, Miktar(*), 2.Miktar,
    //          Öncelik, Açıklama, Revizyon N, Ek Alan-1, Ek Alan-2, Oto. Yarı M,
    //          Oto. Yarı M, Bakiye (0/1), Mamüller Ölçü

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

    const confirmedItems = items.filter((i) => !i.skipped && i.confirmed_stok);

    const rows = confirmedItems.map((item, idx) => ({
      "Fiş No.(*)": `${is_emri_no}-${String(idx + 1).padStart(3, "0")}`,
      Tarih: tarih,
      "Belge Tipi": "ÜSK",
      "İş Emri/Sip.": is_emri_no,
      "Sipariş Kale": "",
      "Depo Öncel": "",
      "L. Depo(*)": "",
      "Çık. Depo(*)": "",
      "*Mamul Kod": item.confirmed_stok?.stok_kodu || "",
      "Fire Depo": "",
      "Miktar(*)": item.confirmed_miktar,
      "2.Miktar": "",
      Öncelik: "",
      Açıklama: item.confirmed_stok?.stok_adi || item.original_urun_adi,
      "Revizyon N": "",
      "Ek Alan-1": "",
      "Ek Alan-2": "",
      "Oto. Yarı M": "",
      "Oto. Yarı M (2)": "",
      "Bakiye (0/1)": "",
      "Mamüller Ölçü": "",
    }));

    // Create workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, { header });

    // Style the header row
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (!ws[addr]) continue;
      ws[addr].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: "1F4E79" } },
        alignment: { horizontal: "center" },
      };
    }

    // Set column widths
    ws["!cols"] = [
      { wch: 18 }, // Fiş No
      { wch: 12 }, // Tarih
      { wch: 12 }, // Belge Tipi
      { wch: 14 }, // İş Emri
      { wch: 12 }, // Sipariş Kale
      { wch: 12 }, // Depo Öncel
      { wch: 10 }, // L. Depo
      { wch: 10 }, // Çık. Depo
      { wch: 18 }, // Mamul Kod
      { wch: 10 }, // Fire Depo
      { wch: 10 }, // Miktar
      { wch: 10 }, // 2.Miktar
      { wch: 10 }, // Öncelik
      { wch: 40 }, // Açıklama
      ...Array(7).fill({ wch: 12 }),
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    // Generate Excel buffer
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const filename = `UretimCikti_${is_emri_no}_${tarih.replace(/\./g, "-")}.xlsx`;

    return new NextResponse(buf, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (error: unknown) {
    console.error("Export error:", error);
    const message =
      error instanceof Error ? error.message : "Export fehlgeschlagen";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
