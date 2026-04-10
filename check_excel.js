const ExcelJS = require('exceljs');
const path = require('path');

async function checkExcel() {
  const file = path.join('d:', 'Users', 'lbh00', 'OneDrive', '문서', '감사', '감사자료', 'audit-automation', 'Braintree', 'task_list_braintree.xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);
  
  workbook.eachSheet((sheet, id) => {
    console.log(`Sheet: ${sheet.name}`);
    const data = [];
    sheet.eachRow((row, rowNum) => {
      data.push(row.values);
    });
    console.log(JSON.stringify(data, null, 2));
  });
}

checkExcel().catch(console.error);
