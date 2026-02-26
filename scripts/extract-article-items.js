// One-time script: read INVENTORY-AS-OF-AUGUST-2026.xlsx and output unique Article/Item values to articles.json
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const excelPath = path.join(__dirname, '..', 'INVENTORY-AS-OF-AUGUST-2026.xlsx');
const outPath = path.join(__dirname, '..', 'Inventory Lab', 'articles.json');

if (!fs.existsSync(excelPath)) {
  console.error('Excel file not found:', excelPath);
  process.exit(1);
}

const workbook = XLSX.readFile(excelPath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

if (!data.length) {
  console.error('Sheet is empty');
  process.exit(1);
}

// Find row that contains "Article/Item" or "Article/It" as header
const colNames = ['Article/Item', 'Article/It'];
let headerRowIndex = -1;
let colIndex = -1;
for (let r = 0; r < Math.min(data.length, 30); r++) {
  const row = data[r].map(h => (h || '').toString().trim());
  for (let c = 0; c < row.length; c++) {
    const cell = row[c];
    if (colNames.some(name => cell === name || cell.replace(/\s/g, '') === name.replace(/\s/g, ''))) {
      headerRowIndex = r;
      colIndex = c;
      break;
    }
  }
  if (headerRowIndex >= 0) break;
}
if (colIndex < 0) {
  colIndex = 0;
  headerRowIndex = 0;
}

// Skip header-like values (column title or common meta text)
const skipValues = new Set([
  'Article/Item', 'Article/It', 'DATE:', 'INVENTORY', 'Laboratory Custodian',
  'OFFICE/CAMPUS:    Multimedia and Speech Laboratory', 'JOVEN T. CRUZ',
  'ICT Equipment, Devices & Accessories', 'ICT EQUIPMENT, DEVICES AND ACCESSORIES',
  'Furniture and Fixtures', 'FURNITURES AND FIXTURES'
]);

const seen = new Set();
const items = [];
const skipUpper = new Set([...skipValues].map(s => s.toUpperCase()));
for (let i = headerRowIndex + 1; i < data.length; i++) {
  const cell = data[i][colIndex];
  const val = (cell != null && String(cell).trim() !== '') ? String(cell).trim() : null;
  if (!val || val.length >= 100) continue;
  const valUpper = val.toUpperCase();
  if (seen.has(valUpper) || skipUpper.has(valUpper)) continue;
  if (valUpper.startsWith('PC USED BY') || valUpper.startsWith('PREPARED BY') || valUpper.startsWith('SUMMARY OF')) continue;
  if (/^PC\s*\d*$/.test(valUpper)) continue; // skip "PC", "PC 1", "PC 2", etc.
  seen.add(valUpper);
  items.push(val); // keep original casing of first occurrence
}

items.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

fs.writeFileSync(outPath, JSON.stringify(items, null, 2), 'utf8');
console.log('Wrote', items.length, 'unique Article/Item values to', outPath);
