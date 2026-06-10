export interface StokItem {
  stok_kodu: string;
  stok_adi: string;
  uretim?: string;
  cesit?: string;
}

export interface ScannedItem {
  urun_adi: string;
  miktar: string;
}

export interface ConfirmedItem {
  original_urun_adi: string;
  original_miktar: string;
  confirmed_stok: StokItem | null;
  confirmed_miktar: number;
  skipped: boolean;
}

export interface ScanResult {
  is_emri_no: string;
  tarih: string;
  sira_no: string;
  items: ScannedItem[];
}

export interface ExportPayload {
  is_emri_no: string;
  tarih: string;
  items: ConfirmedItem[];
}
