const { chromium } = require('playwright');
const ExcelJS = require('exceljs');

/**
 * 브라우저 초기화 및 페이지 반환
 * @param {boolean} headless 헤드리스 모드 여부
 * @returns {Promise<{browser: import('playwright').Browser, page: import('playwright').Page}>}
 */
async function initBrowser(headless = true) {
    const browser = await chromium.launch({ headless });
    const context = await browser.newContext();
    const page = await context.newPage();
    return { browser, page };
}

/**
 * 공통 엑셀 파일 로드
 * @param {string} filePath 엑셀 파일 경로
 * @returns {Promise<import('exceljs').Workbook>}
 */
async function loadExcelTemplate(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    return workbook;
}

/**
 * 엑셀 파일 저장
 * @param {import('exceljs').Workbook} workbook 
 * @param {string} filePath 
 */
async function saveExcelFile(workbook, filePath) {
    await workbook.xlsx.writeFile(filePath);
}

module.exports = {
    initBrowser,
    loadExcelTemplate,
    saveExcelFile
};
