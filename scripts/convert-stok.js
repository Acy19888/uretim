/**
 * Einmaliges Script: Stok Kart Kayıtları.xlsx → lib/stokData.json
 * Ausführen mit: node scripts/convert-stok.js
 */
const XLSX = require('xlsx');
const fs   = require('fs');
const path = require('path');

const excelPath = path.join(__dirname, '..', 'Stok Kart Kayıtları.xlsx');
const outPath   = path.join(__dirname, '..', 'lib', 'stokData.json');

if (!fs.existsSync(excelPath)) {
  console.error('❌ Datei nicht gefunden:', excelPath);
  process.exit(1);
}

const wb   = XLSX.readFile(excelPath);
const ws   = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { raw: false, defval: '' });

const items = [];
for (const row of rows) {
  const kodu  = (row['STOK_KODU'] || row['Stok Kodu'] || '').trim();
  const adi   = (row['STOK_ADI']  || row['Stok Adı']  || '').trim();
  const cesit = (row['Çeşit']     || '').trim();
  if (kodu && adi) items.push({ stok_kodu: kodu, stok_adi: adi, cesit });
}

fs.writeFileSync(outPath, JSON.stringify(items));
console.log(`✅ ${items.length} Produkte gespeichert → lib/stokData.json`);
