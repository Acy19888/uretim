import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Üretim Scanner",
  description: "Windoform Üretim Çıktı Takip Sistemi",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="tr">
      <body className="min-h-screen bg-slate-100">
        <header className="bg-blue-800 text-white px-4 py-3 flex items-center gap-3 shadow-lg">
          <div className="text-2xl font-bold tracking-tight">WINDOFORM</div>
          <div className="text-sm opacity-80 border-l border-blue-600 pl-3">
            Üretim Çıktı Tarayıcı
          </div>
        </header>
        <main className="max-w-2xl mx-auto px-4 py-6">{children}</main>
      </body>
    </html>
  );
}
