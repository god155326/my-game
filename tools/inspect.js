const XLSX = require('/tmp/xlsxbuild/node_modules/xlsx');
const fs = require('fs');
const wb = XLSX.readFile('data/不朽之旅.xlsx');
const name = process.argv[2];
const ws = wb.Sheets[name];
if (!ws) { console.log('NO SHEET', name, 'available:', wb.SheetNames); process.exit(1); }
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
console.log(JSON.stringify(rows.slice(0, Number(process.argv[3]||8)), null, 1));
