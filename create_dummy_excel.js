const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function createDummyExcel(companyFolder) {
    const folderPath = path.join(__dirname, companyFolder);
    if (!fs.existsSync(folderPath)) {
        fs.mkdirSync(folderPath, { recursive: true });
    }
    
    const filePath = path.join(folderPath, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    
    worksheet.columns = [
        { header: 'ID', key: 'id', width: 10 },
        { header: 'Amount', key: 'amount', width: 20 },
        { header: 'Status', key: 'status', width: 15 }
    ];
    
    await workbook.xlsx.writeFile(filePath);
    console.log(`Created template.xlsx in ${companyFolder}`);
}

async function main() {
    await createDummyExcel('Company_A');
    await createDummyExcel('Company_B');
    await createDummyExcel('Braintree');
}

main().catch(console.error);
