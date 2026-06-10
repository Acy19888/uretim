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
        <header
          className="text-white px-5 py-3.5 shadow-lg flex items-center gap-3"
          style={{ background: "var(--wf-blue)" }}
        >
          {/* WINDOFORM W-mark */}
          <svg width="32" height="27" viewBox="0 0 56 48" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M0,0 L9,48 L28,12 L47,48 L56,0 L49,0 L28,36 L7,0 Z" fill="white"/>
            <path d="M9,48 L14,48 L28,24 L42,48 L47,48 L28,12 Z" fill="#2B5597"/>
          </svg>
          <div className="flex items-baseline gap-0.5">
            <span className="text-xl font-black" style={{ letterSpacing: "0.12em" }}>WINDOFORM</span>
            <span className="text-[9px] align-super opacity-60 ml-0.5">®</span>
          </div>
          <div className="h-4 w-px bg-white/30 mx-1" />
          <span className="text-sm opacity-70 font-medium">Üretim Tarayıcı</span>
        </header>
        <main className="max-w-xl mx-auto px-4 py-5">{children}</main>
      </body>
    </html>
  );
}
