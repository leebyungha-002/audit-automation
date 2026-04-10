const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { initBrowser, loadExcelTemplate, saveExcelFile } = require('./index');

/**
 * 파일 다운로드 및 저장 통합 처리 헬퍼
 */
async function handleDownloadAndSave(page, downloadBtnSelector, targetName, rawDataDir, menuName, filePrefix = '') {
    console.log(`[${menuName}] 결과 다운로드 버튼 노출 대기중...`);
    await page.waitForSelector(downloadBtnSelector, { state: 'visible', timeout: 30000 });
    
    console.log(`[${menuName}] 다운로드를 진행합니다.`);
    const downloadPromise = page.waitForEvent('download');
    await page.click(downloadBtnSelector);
    
    const download = await downloadPromise;
    const downloadPath = await download.path();
    console.log(`[${menuName}] 임시 다운로드 캡처 완료: ${downloadPath}`);
    
    // 마스터 단일 파일 병합 대상
    if (menuName === '상세 거래 검색' || menuName === '총계정원장조회') {
        const baseMasterFileName = menuName === '상세 거래 검색' ? '상세거래검색.xlsx' : '총계정원장.xlsx';
        const masterFileName = `${filePrefix}${baseMasterFileName}`;
        const masterPath = path.join(rawDataDir, masterFileName);
        
        console.log(`[${menuName}] 마스터 파일(${masterFileName})에 '${targetName}' 시트로 병합 처리...`);
        
        const masterBook = new ExcelJS.Workbook();
        if (fs.existsSync(masterPath)) {
            await masterBook.xlsx.readFile(masterPath);
        }
        
        const srcBook = new ExcelJS.Workbook();
        await srcBook.xlsx.readFile(downloadPath);
        const srcSheet = srcBook.worksheets[0];
        
        // 시트명 31자 제한 안전 처리 및 특수문자 제거
        const safeSheetName = targetName.substring(0, 31).replace(/[\\/?*[\]]/g, '_');
        let destSheet = masterBook.getWorksheet(safeSheetName);
        if (destSheet) {
            masterBook.removeWorksheet(safeSheetName);
        }
        destSheet = masterBook.addWorksheet(safeSheetName);
        
        srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const destRow = destSheet.getRow(rowNumber);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                destRow.getCell(colNumber).value = cell.value;
            });
        });
        
        await masterBook.xlsx.writeFile(masterPath);
        console.log(`[${menuName}] 마스터 파일 병합 완료.`);
    } else {
        // 개별 파일 저장 대상
        const finalTargetName = targetName.startsWith(filePrefix) ? targetName : `${filePrefix}${targetName}`;
        const finalPath = path.join(rawDataDir, finalTargetName + (finalTargetName.endsWith('.xlsx') ? '' : '.xlsx'));
        fs.copyFileSync(downloadPath, finalPath);
        console.log(`[${menuName}] 개별 파일 저장 완료: ${finalPath}`);
    }
}

/**
 * 범용 재무감사 자동화 러너
 * @param {Object} config - 실행할 회사의 설정 객체 (config.js)
 * @param {string} companyDir - 실행할 회사의 절대 경로 (__dirname)
 */
async function runAudit(config, companyDir) {
    const companyName = config.companyName || path.basename(companyDir);
    console.log(`=== ${companyName} 감사 자동화 시작 ===`);
    
    // 설정에서 Debug 여부 확인
    const isHeadless = config.taskList?.RunMode === 'Debug' ? false : true;
    
    const clientName = config.taskList?.ClientName || config.companyName || companyName;
    const targetYear = config.taskList?.TargetYear || '';
    const filePrefix = targetYear ? `${clientName}_${targetYear}_` : `${clientName}_`;
    
    const { browser, page } = await initBrowser(isHeadless); 
    
    try {
        // Settings 시트(config.taskList.Url)에서 우선적으로 가져오며, 그 외에는 config.url로 우회
        const url = config.taskList?.Url || config.url;
        if (!url || url.includes('127.0.0.1')) {
            console.log(`[경고] 연결 주소가 로컬 호스트로 설정되어 있을 수 있습니다. 실제 웹사이트 URL을 입력해주세요.`);
        }

        console.log(`[${companyName}] 접속 시도 중인 URL: ${url}`);
        // 페이지 로딩 완료 기준을 networkidle로 바인딩하여 60초까지 대기
        await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });
        
        // --- 0. 로그인 로직 ---
        if (config.credentials && config.credentials.userId) {
            console.log(`[${companyName}] 로그인을 시도합니다...`);
            const emailSelector = config.selectors.loginId || 'input[type="email"]';
            const pwSelector = config.selectors.loginPassword || 'input[type="password"]';
            const loginBtnSelector = config.selectors.loginButton || 'button:has-text("로그인")';
            
            // 로그인 입력 요소가 보일 때까지 대기
            await page.waitForSelector(emailSelector, { state: 'visible', timeout: 60000 });
            
            // 데이터 입력
            await page.fill(emailSelector, config.credentials.userId || '');
            await page.fill(pwSelector, config.credentials.userPassword || '');
            
            // 로그인 버튼 클릭
            console.log(`[${companyName}] 로그인 버튼을 클릭합니다.`);
            await page.click(loginBtnSelector);
            console.log(`[${companyName}] 화면 전환 또는 로딩 대기...`);
        } else {
            console.log(`[${companyName}] 로그인 인증 정보가 제공되지 않아 로그인을 건너뜁니다.`);
        }
        
        // --- 1. 파일 업로드 로직 ---
        if (config.uploadFileName) {
            console.log(`[${companyName}] 파일 선택: 업로드를 진행합니다...`);
            const uploadFilePath = path.join(companyDir, config.uploadFileName);
            const fileSelector = config.selectors.fileUploadInput || 'input[type="file"]';
            
            console.log(`[${companyName}] 화면에서 업로드 버튼(${fileSelector})을 대기 중 (최대 60초 대기)`);
            await page.waitForSelector(fileSelector, { state: 'attached', timeout: 60000 });
            
            // 지정된 로컬 경로의 파일을 업로드
            await page.setInputFiles(fileSelector, uploadFilePath);
            console.log(`[${companyName}] 파일 처리 대기 중...`);
        }
        
        // --- 2. 메뉴 순회 및 동작 로직 ---
        const rawDataDir = path.join(companyDir, 'raw_data');
        if (!fs.existsSync(rawDataDir)) {
            fs.mkdirSync(rawDataDir, { recursive: true });
        }

        if (!config.menus || config.menus.length === 0) {
            console.log(`[${companyName}] 실행할 메뉴(지시서 시트)가 없습니다. 자동화를 종료합니다.`);
            return;
        }

        for (const menu of config.menus) {
            const menuName = menu.menuName;
            console.log(`\n=== [메뉴 진입] ${menuName} ===`);
            
            // 시트 이름에 해당하는 메뉴(버튼/링크)로 화면 이동
            const menuSelector = `text="${menuName}"`;
            const menuHandle = await page.waitForSelector(menuSelector, { state: 'visible', timeout: 30000 });
            
            // 새 창(새 탭) 열림을 방지하기 위해 target="_blank" 속성 강제 제거
            await menuHandle.evaluate(node => {
                if (node.hasAttribute('target')) {
                    node.removeAttribute('target');
                }
                // 부모 a 태그가 있을 경우도 처리
                const parentA = node.closest('a');
                if (parentA && parentA.hasAttribute('target')) {
                    parentA.removeAttribute('target');
                }
            });
            await menuHandle.click();
            
            // 안정적 전환을 위한 명시적 대기
            await page.waitForTimeout(2000);
            
            if (menuName === '상세 거래 검색' || menuName === '총계정원장조회' || menuName === '벤포드 법칙 분석') {
                for (const task of menu.tasks) {
                    // 항목 추출. '총계정원장조회'와 같이 계정명이 A열인 경우 객체의 첫번째 키 활용
                    const taskKeys = Object.keys(task);
                    if (taskKeys.length === 0) continue;
                    
                    const accountName = task['계정과목'] || task[taskKeys[0]];
                    if (!accountName) {
                        console.log(`[${menuName}] 계정과목(또는 첫 번째 열) 값이 없어 건너뜁니다:`, task);
                        continue;
                    }
                    console.log(`\n--- [${accountName}] 처리 시작 ---`);
                    
                    const comboboxSelector = config.selectors.accountCombobox || 'button[role="combobox"]';
                    
                    // 초기화 버튼 처리
                    if (config.selectors.resetButton) {
                        try {
                            await page.click(config.selectors.resetButton, { timeout: 2000 });
                        } catch(e) { /* 초기화 버튼 노출안됨 무시 */ }
                    }
                    
                    // 계정과목 입력 및 선택
                    await page.waitForSelector(comboboxSelector, { state: 'visible' });
                    await page.click(comboboxSelector);
                    await page.waitForTimeout(500); 
                    
                    // 기존 내용 지우기
                    await page.keyboard.press('Control+A'); // Windows/Linux
                    await page.keyboard.press('Meta+A'); // Mac
                    await page.keyboard.press('Backspace');
                    await page.waitForTimeout(300);
                    
                    // 입력 및 선택 엔터
                    await page.keyboard.type(accountName, { delay: 50 });
                    await page.waitForTimeout(500);
                    await page.keyboard.press('Enter');
                    await page.waitForTimeout(500);
                    
                    if (menuName === '총계정원장조회') {
                        // 검색 버튼 없음, 테이블 노출까지 대기
                        console.log(`[${accountName}] 정보 입력 완료, 테이블 갱신을 위해 2~3초 대기합니다.`);
                        await page.waitForTimeout(3000); // 3초 대기
                        
                        const downloadBtnSelector = 'button:has-text("엑셀 다운로드")';
                        await handleDownloadAndSave(page, downloadBtnSelector, accountName, rawDataDir, menuName, filePrefix);

                    } else {
                        // 상세 거래 검색, 벤포드 법칙 분석
                        if (menuName === '상세 거래 검색' && task['표시방식']) {
                            const rbLabel = task['표시방식'];
                            console.log(`[${accountName}] 라디오 버튼 선택: ${rbLabel}`);
                            try {
                                await page.locator(`label:has-text("${rbLabel}")`).click({ timeout: 5000 });
                                await page.waitForTimeout(500);
                            } catch(e) {
                                console.log(`[경고] 라디오 버튼 ${rbLabel} 요소를 찾을 수 없습니다.`);
                            }
                        }
                        
                        // 공통 검색 버튼 클릭
                        const searchBtnSelector = config.selectors.searchButton || 'button:has-text("검색")';
                        await page.click(searchBtnSelector);
                        
                        // 결과 테이블 및 다운로드 대기
                        await page.waitForTimeout(1000); // 검색 반응 딜레이
                        
                        const downloadBtnSelector = config.selectors.excelDownloadBtn || 'button:has-text("결과 다운로드")';
                        
                        // 상세 거래 검색은 단일 파일 병합 시 시트명을 계정명으로 하므로 accountName 전달. 
                        // 개별 파일 저장인 경우 파일명을 전달.
                        const targetName = (menuName === '상세 거래 검색') ? accountName : (task['파일명'] || accountName);
                        await handleDownloadAndSave(page, downloadBtnSelector, targetName, rawDataDir, menuName, filePrefix);
                    }
                }

            } else if (menuName === '매입/매출 이중거래처 분석') {
                console.log(`\n--- [${menuName}] 처리 시작 ---`);
                const task = menu.tasks && menu.tasks.length > 0 ? menu.tasks[0] : {};
                
                await page.click('button:has-text("이중거래처 분석 시작")');
                
                await page.waitForTimeout(1000);
                const downloadBtnSelector = 'button:has-text("결과 다운로드")';
                
                const fileName = task['파일명'] || '이중거래처_결과';
                await handleDownloadAndSave(page, downloadBtnSelector, fileName, rawDataDir, menuName, filePrefix);
            } else {
                console.log(`[${menuName}] 현재 구현되지 않은 메뉴 형식입니다. 생략합니다.`);
            }
        }
        
        console.log(`=== ${companyName} 자동화 종료 ===`);
        
    } catch (error) {
        console.error(`[${companyName}] 실행 중 오류 발생:`, error);
        
        try {
            const screenshotPath = path.join(companyDir, 'error.png');
            await page.screenshot({ path: screenshotPath, fullPage: true });
            console.log(`[${companyName}] 에러 스크린샷 캡쳐 완료: ${screenshotPath}`);
        } catch (captureError) {
            console.error(`[${companyName}] 스크린샷 캡쳐 실패:`, captureError);
        }
    } finally {
        await browser.close();
    }
}

module.exports = runAudit;
