# Windoform Üretim Scanner

Handgeschriebene Produktionszettel fotografieren → KI erkennt die Schrift → Zeile für Zeile bestätigen → Excel-Export.

## Features

- 📷 Kamera-Scan (mobil + Desktop)
- 🤖 OpenAI GPT-4 Vision liest Handschrift
- 🔍 Automatischer Produktabgleich mit Stok Kart Kayıtları
- ✅ Zeile-für-Zeile-Bestätigung mit Ja/Nein
- 📥 Excel-Export im boş Üsk Format

## Setup

### 1. Repository klonen & Pakete installieren

```bash
git clone https://github.com/Acy19888/uretim.git
cd uretim
npm install
```

### 2. Lokale Entwicklung

```bash
cp .env.example .env.local
# .env.local öffnen und OPENAI_API_KEY eintragen
npm run dev
```

### 3. Vercel Deployment

1. GitHub repo mit Vercel verbinden: https://vercel.com/new
2. In Vercel → Settings → Environment Variables:
   - `OPENAI_API_KEY` = dein OpenAI API Key
3. Deploy!

## Erste Nutzung

1. App öffnen
2. `Stok Kart Kayıtları.xlsx` hochladen (einmalig, wird im Browser gespeichert)
3. Produktionszettel fotografieren
4. Jede Zeile bestätigen oder korrigieren
5. Excel herunterladen
