const ExcelJS = require('exceljs');
const paths = [
  'C:/Users/dell/Desktop/Employee_KPI_Report_20261102.xlsx',
  'F:/3 ALPHA/Call report sample (Oct 2026).xlsx'
];

async function dump(path) {
  const wb = new ExcelJS.Workbook();
  console.log('\n=== FILE:', path, '===');
  try {
    await wb.xlsx.readFile(path);
    console.log('SHEETS:', wb.worksheets.map(s => s.name));
    for (const ws of wb.worksheets) {
      console.log('\n--- SHEET', ws.name, 'rows=', ws.rowCount, 'cols=', ws.columnCount);
      for (let r = 1; r <= Math.min(ws.rowCount, 12); r++) {
        const row = ws.getRow(r).values;
        console.log(JSON.stringify(row));
      }
    }
  } catch (e) {
    console.log('ERROR:', e && e.message ? e.message : e);
  }
}

(async () => {
  for (const p of paths) await dump(p);
})();
