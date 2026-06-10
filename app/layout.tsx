import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "WINDOFORM | Üretim Tarayıcı",
  description: "Yarı Mamul Giriş ve Çıkış Takip Sistemi",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="tr">
      <body className="min-h-screen" style={{ background: "var(--wf-bg)" }}>
        <header className="bg-white shadow-md border-b-4 flex items-center justify-between px-5 py-3"
          style={{ borderBottomColor: "var(--wf-blue)" }}>
          {/* Real WINDOFORM logo */}
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src="/windoform-logo.png" alt="WINDOFORM" className="h-9 w-auto object-contain" />
          <span className="text-xs font-semibold uppercase tracking-widest"
            style={{ color: "var(--wf-gray)" }}>
            Üretim Tarayıcı
          </span>
        </header>
        <main className="max-w-xl mx-auto px-4 py-5">{children}</main>
      </body>
    </html>
  );
}
