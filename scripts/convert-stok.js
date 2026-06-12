/**
 * Tek seferlik script: Stok Kart Kayıtları.xlsx → lib/stokData.json
 * Çalıştır: node scripts/convert-stok.js
 */
const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const excelPath = path.join(__dirname, '..', 'Stok Kart Kayıtları.xlsx');
const outPath   = path.join(__dirname, '..', 'lib', 'stokData.json');

if (!fs.existsSync(excelPath)) {
  console.error('❌ Dosya bulunamadı:', excelPath);
  process.exit(1);
}

const wb   = XLSX.readFile(excelPath);
const ws   = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { raw: false, defval: '' });

const items = [];
let atlanan = 0;
for (const row of rows) {
  const kodu   = (row['STOK_KODU'] || row['Stok Kodu'] || '').trim();
  const adi    = (row['STOK_ADI']  || row['Stok Adı']  || '').trim();
  const cesit  = (row['Çeşit']     || '').trim();
  const uretim = (row['Üretim']    || row['URETIM']    || row['Uretim'] || '').trim().toUpperCase();

  if (!kodu || !adi) continue;

  // Sadece YARI MAMUL — 2K ve MAMUL atla
  if (uretim && uretim !== 'YARI MAMUL') { atlanan++; continue; }

  items.push({ stok_kodu: kodu, stok_adi: adi, cesit });
}

fs.writeFileSync(outPath, JSON.stringify(items));
console.log(`✅ ${items.length} YARI MAMUL ürün kaydedildi → lib/stokData.json`);
if (atlanan > 0) console.log(`⏭  ${atlanan} ürün atlandı (2K / MAMUL)`);
