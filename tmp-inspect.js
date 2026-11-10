const ExcelJS = require('exceljs');
const path = 'C:/Users/dell/Desktop/testing folder/tset 3.xlsx';
const wb = new ExcelJS.Workbook();
wb.xlsx.readFile(path)
  .then(() => {
    console.log('SHEETS', wb.worksheets.map((s) => s.name));
    for (const ws of wb.worksheets) {
      console.log('---', ws.name, 'rows', ws.rowCount, 'cols', ws.columnCount);
      for (let r = 1; r <= Math.min(ws.rowCount, 12); r++) {
        const row = ws.getRow(r).values;
        console.log(JSON.stringify(row));
      }
    }
  })
  .catch((e) => {
    console.error(e);
    process.exit(1);
  });
