import type { ExportPayload } from "./types";

export interface OneDriveUploadPayload {
  exportPayload: ExportPayload;
  photos: { data: string; mime: string }[];
  baseName: string;
}

/**
 * Excel + fotoğrafları sunucu üzerinden OneDrive'a yükler.
 * Telefonda login gerekmez — kimlik bilgileri Vercel'de saklanır.
 */
export async function uploadToOneDrive(
  payload: OneDriveUploadPayload
): Promise<{ uploaded: number; folder: string }> {
  const res = await fetch("/api/onedrive", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "OneDrive yükleme başarısız");
  return data;
}
