const ExcelJS = require('exceljs');
const fs = require('fs');

const source = 'c:\\Users\\dell\\Desktop\\Employee_KPI_Report_2026 11 06.xlsx';
const generated = 'c:\\Users\\dell\\Desktop\\testing folder\\2 (Oct 2026)-kpi.xlsx';

(async () => {
  console.log('source exists:', fs.existsSync(source));
  console.log('generated exists:', fs.existsSync(generated));

  if (!fs.existsSync(source) || !fs.existsSync(generated)) {
    return;
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(source);
  const ws = wb.worksheets[0];
  const header = [];
  for (let c = 1; c <= Math.min(ws.columnCount, 80); c++) {
    const value = ws.getRow(1).getCell(c).value;
    header.push(value == null ? '' : String(value).toLowerCase());
  }

  const nameIdx = header.findIndex((h) => h.includes('employee name') || h.includes('name'));
  const codeIdx = header.findIndex((h) => h.includes('emp. id') || h.includes('emp id') || h.includes('employee id') || h.includes('id'));

  console.log('nameIdx', nameIdx, 'codeIdx', codeIdx);

  const sourceRows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const nameCell = nameIdx >= 0 ? row.getCell(nameIdx + 1).value : null;
    const codeCell = codeIdx >= 0 ? row.getCell(codeIdx + 1).value : null;
    const name = String(nameCell ?? '').trim();
    const code = String(codeCell ?? '').trim();
    if (name || code) {
      sourceRows.push({ name, code });
    }
  }

  console.log('source rows:', sourceRows.length);
  console.log('first 10 source rows:', sourceRows.slice(0, 10));

  const byName = new Map();
  for (const row of sourceRows) {
    const key = row.name.toLowerCase();
    byName.set(key, (byName.get(key) || 0) + 1);
  }
  const dupNames = [...byName.entries()].filter(([, count]) => count > 1).slice(0, 20);
  console.log('duplicate source names:', dupNames.length);
  console.log('sample duplicate names:', dupNames.slice(0, 10));

  const generatedWb = new ExcelJS.Workbook();
  await generatedWb.xlsx.readFile(generated);
  const genWs = generatedWb.worksheets[0];
  const generatedRows = [];
  for (let r = 2; r <= genWs.rowCount; r++) {
    const row = genWs.getRow(r);
    const code = row.getCell(1).value;
    const name = row.getCell(2).value;
    const total = row.getCell(9).value;
    if (code || name) {
      generatedRows.push({ code: String(code ?? '').trim(), name: String(name ?? '').trim(), total: total ?? 0 });
    }
  }

  console.log('generated rows:', generatedRows.length);
  console.log('first 10 generated rows:', generatedRows.slice(0, 10));
})();
