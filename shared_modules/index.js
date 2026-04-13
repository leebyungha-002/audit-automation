'use strict';
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');

/**
 * 브라우저 초기화 및 페이지 반환
 * @param {boolean} headless 헤드리스 모드 여부
 * @param {string|null} persistentProfile 로그인 세션을 유지할 로컬 프로필 디렉토리 경로.
 *   null이면 매번 새 세션(Incognito)으로 실행.
 *   경로를 지정하면 쿠키/세션이 해당 폴더에 저장되어 다음 실행 시 자동 로그인됨.
 * @returns {Promise<{browser: object, page: import('playwright').Page}>}
 */
async function initBrowser(headless = true, persistentProfile = null) {
    if (persistentProfile) {
        // launchPersistentContext는 BrowserContext를 직접 반환함 (Browser 아님)
        // BrowserContext도 .close()를 갖고 있으므로 browser 자리에 넣어 사용 가능
        const context = await chromium.launchPersistentContext(persistentProfile, {
            headless,
            args: ['--no-sandbox', '--disable-setuid-sandbox'],
        });
        const page = await context.newPage();
        console.log(`[브라우저] 세션 유지 모드로 실행됩니다. (프로필: ${persistentProfile})`);
        return { browser: context, page };
    }

    const browser = await chromium.launch({ headless });
    const context = await browser.newContext();
    const page = await context.newPage();
    console.log('[브라우저] 새 세션(Incognito)으로 실행됩니다.');
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
    saveExcelFile,
};
