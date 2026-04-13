'use strict';
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { initBrowser } = require('./index');

// ─── 엔드포인트 라우팅 ────────────────────────────────────────────────────────
const DEFAULT_ENDPOINT_MAP = {
    '분개장':      '/ai-analysis',
    '분개장분석':  '/ai-analysis',
    '분개장 분석': '/ai-analysis',
};

function getMenuEndpoint(menuName, config) {
    return config.menuEndpoints?.[menuName]
        ?? DEFAULT_ENDPOINT_MAP[menuName]
        ?? '/analysis';
}

// ─── 시트명 → UI 카드/버튼 텍스트 매핑 ──────────────────────────────────────
// task_list.xlsx의 시트명(menuName)과 웹 앱의 실제 카드 텍스트가 다를 때 사용.
// config.menuLabels 로 회사별 커스텀 오버라이드 가능.
const DEFAULT_MENU_LABEL_MAP = {
    '총계정원장':             '총계정원장 조회',
    '상세거래검색':           '상세 거래 검색',
    '이중거래처분석':         '매입/매출 이중거래처 분석',
    '벤포드':                 '벤포드 법칙 분석',
    '계정연관거래처':         '계정 연관 거래처 분석',
    '외상매출매입상계':       '외상매출/매입 상계 거래처 분석',
    '추정손익':               '추정 손익 분석',
    '매출관리비추이':         '매출/관리비 월별 추이 분석',
    '전기비교':               '전기 데이터 비교 분석',
    '감사샘플링':             '감사 샘플링',
    '금감원위험분석':         '금감원 지적사례 기반 위험 분석',
    '재무제표증감':           '재무제표 증감 분석',
};

function getMenuUiLabel(menuName, config) {
    return config.menuLabels?.[menuName]
        ?? DEFAULT_MENU_LABEL_MAP[menuName]
        ?? menuName; // 매핑 없으면 시트명 그대로 사용
}

// ─── 전기(전년도) 계정별원장 파일 탐색 ──────────────────────────────────────────
// raw_data 폴더에서 '전기' 또는 전년도 연도(현재연도-1)가 포함된 xlsx 파일을 반환.
function findPrevYearLedgerFile(rawDataDir) {
    const prevYear = String(new Date().getFullYear() - 1);
    if (!fs.existsSync(rawDataDir)) return null;

    const found = fs.readdirSync(rawDataDir).find(f => {
        const ext = path.extname(f).toLowerCase();
        return (f.includes('전기') || f.includes(prevYear)) && (ext === '.xlsx' || ext === '.xls');
    });

    return found ? path.join(rawDataDir, found) : null;
}

// ─── /analysis 엔드포인트: 계정별원장 파일 업로드 ─────────────────────────────
// 업로드 UI가 보일 때만 실행. 이미 분석 화면이면 건너뜀.
async function uploadGeneralLedgerIfNeeded(page, config, companyDir) {
    const fileInputSelector = config.selectors.fileUploadInput || 'input[type="file"]';
    const uploadAreaExists = await page.$(fileInputSelector).then(el => !!el).catch(() => false);

    if (!uploadAreaExists) {
        console.log('[업로드] 업로드 영역이 없습니다. 이미 분석 화면이거나 파일이 불필요합니다.');
        return;
    }

    const uploadRelPath = config.uploadFileName;
    if (!uploadRelPath) {
        console.log('[업로드] config.uploadFileName이 설정되지 않았습니다.');
        return;
    }

    const uploadFilePath = path.join(companyDir, uploadRelPath);
    if (!fs.existsSync(uploadFilePath)) {
        console.log(`[업로드] 파일이 존재하지 않습니다: ${uploadFilePath}`);
        return;
    }

    // ── 전기 파일 사전 스캔 ──────────────────────────────────────────────────
    const rawDataDir = path.join(companyDir, 'raw_data');
    const prevYearFile = findPrevYearLedgerFile(rawDataDir);
    if (prevYearFile) {
        console.log(`[업로드] 전기 데이터 발견: ${path.basename(prevYearFile)} 업로드를 시작합니다.`);
    } else {
        console.log('[업로드] 전기 데이터 미발견: 당기 분석만 진행합니다.');
    }

    console.log(`[업로드] 계정별원장 파일 업로드 시작: ${path.basename(uploadFilePath)}`);
    await page.setInputFiles(fileInputSelector, uploadFilePath);

    // 업로드 후 처리 완료 대기: 업로드 영역이 사라지거나 분석 UI가 나타날 때까지
    console.log('[업로드] 파일 처리 대기 중...');
    try {
        await page.waitForSelector(fileInputSelector, { state: 'hidden', timeout: 30000 });
        console.log('[업로드] 처리 완료. 분석 화면으로 전환됨.');
    } catch {
        console.log('[업로드] 전환 감지 실패. 5초 추가 대기...');
        await page.waitForTimeout(5000);
    }

    // ── 전기 데이터 업로드 여부 다이얼로그 처리 ─────────────────────────────
    if (prevYearFile) {
        // 전기 파일이 있을 때: "네, 전기 데이터도 업로드하겠습니다" 클릭 후 파일 주입
        try {
            const yesPrevBtn = await page.waitForSelector(
                'button:has-text("네")',
                { state: 'visible', timeout: 5000 }
            );
            await yesPrevBtn.click();
            await page.waitForTimeout(1000);

            // 전기 업로드용 file input 대기 후 주입 (새로 나타난 마지막 input 사용)
            await page.waitForSelector('input[type="file"]', { state: 'attached', timeout: 10000 });
            const allInputs = await page.locator('input[type="file"]').all();
            const prevInput = allInputs[allInputs.length - 1];
            await prevInput.setInputFiles(prevYearFile);
            console.log(`[업로드] 전기 파일 주입 완료: ${path.basename(prevYearFile)}`);
            await page.waitForTimeout(1000);
        } catch (e) {
            console.log(`[업로드] 전기 데이터 다이얼로그 처리 실패 (건너뜀): ${e.message}`);
        }
    } else {
        // 전기 파일이 없을 때: "아니요, 당기만 분석하겠습니다" 클릭
        try {
            const skipPrevBtn = await page.waitForSelector(
                'button:has-text("아니요")',
                { state: 'visible', timeout: 5000 }
            );
            console.log('[업로드] 전기 데이터 업로드 다이얼로그 감지. 당기만 분석으로 진행합니다.');
            await skipPrevBtn.click();
            await page.waitForTimeout(1000);
        } catch {
            // 다이얼로그가 없으면 그냥 진행
        }
    }
}

// ─── 다운로드 & 저장 헬퍼 ────────────────────────────────────────────────────
async function handleDownloadAndSave(page, downloadBtnSelector, targetName, rawDataDir, menuName, filePrefix = '') {
    console.log(`[${menuName}] 결과 다운로드 버튼 대기 중...`);
    await page.waitForSelector(downloadBtnSelector, { state: 'visible', timeout: 30000 });

    console.log(`[${menuName}] 다운로드를 진행합니다.`);
    const downloadPromise = page.waitForEvent('download');
    await page.click(downloadBtnSelector);
    const download = await downloadPromise;
    const downloadPath = await download.path();
    console.log(`[${menuName}] 임시 다운로드 캡처 완료.`);

    // 마스터 파일 병합 대상
    const MASTER_MERGE_MENUS = ['상세 거래 검색', '총계정원장조회', '총계정원장'];
    if (MASTER_MERGE_MENUS.includes(menuName)) {
        const baseFileName = (menuName === '상세 거래 검색') ? '상세거래검색.xlsx' : '총계정원장.xlsx';
        const masterPath = path.join(rawDataDir, `${filePrefix}${baseFileName}`);

        const masterBook = new ExcelJS.Workbook();
        if (fs.existsSync(masterPath)) await masterBook.xlsx.readFile(masterPath);

        const srcBook = new ExcelJS.Workbook();
        await srcBook.xlsx.readFile(downloadPath);
        const srcSheet = srcBook.worksheets[0];

        const safeSheetName = targetName.substring(0, 31).replace(/[\\/?*[\]]/g, '_');
        if (masterBook.getWorksheet(safeSheetName)) masterBook.removeWorksheet(safeSheetName);
        const destSheet = masterBook.addWorksheet(safeSheetName);

        srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const destRow = destSheet.getRow(rowNumber);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                destRow.getCell(colNumber).value = cell.value;
            });
        });

        await masterBook.xlsx.writeFile(masterPath);
        console.log(`[${menuName}] 마스터 파일에 '${targetName}' 시트 병합 완료.`);
    } else {
        const finalName = targetName.startsWith(filePrefix) ? targetName : `${filePrefix}${targetName}`;
        const finalPath = path.join(rawDataDir, finalName.endsWith('.xlsx') ? finalName : `${finalName}.xlsx`);
        fs.copyFileSync(downloadPath, finalPath);
        console.log(`[${menuName}] 개별 파일 저장 완료: ${path.basename(finalPath)}`);
    }
}

// ─── /analysis 엔드포인트 메뉴 핸들러 ────────────────────────────────────────
// 시트명("총계정원장", "총계정원장조회" 등)을 모두 처리.
async function handleAnalysisMenu(page, menu, config, rawDataDir, filePrefix) {
    const { menuName, tasks } = menu;

    // "총계정원장" 계열: 계정 콤보박스 → 대기 → 다운로드
    const IS_LEDGER_MENU = ['총계정원장', '총계정원장조회'].includes(menuName);
    // "상세 거래 검색" / "벤포드 법칙 분석": 계정 선택 + 검색 버튼 → 다운로드
    const IS_SEARCH_MENU = ['상세 거래 검색', '벤포드 법칙 분석'].includes(menuName);

    if (IS_LEDGER_MENU || IS_SEARCH_MENU) {
        for (const task of tasks) {
            const taskKeys = Object.keys(task);
            if (taskKeys.length === 0) continue;

            const accountName = String(task['계정과목'] ?? task[taskKeys[0]] ?? '');
            if (!accountName) {
                console.log(`[${menuName}] 계정과목 값이 없어 건너뜁니다:`, task);
                continue;
            }
            console.log(`\n--- [${accountName}] 처리 시작 ---`);

            const comboboxSelector = config.selectors.accountCombobox || 'button[role="combobox"]';

            if (config.selectors.resetButton) {
                try { await page.click(config.selectors.resetButton, { timeout: 2000 }); }
                catch { /* 초기화 버튼 없으면 무시 */ }
            }

            await page.waitForSelector(comboboxSelector, { state: 'visible' });
            await page.click(comboboxSelector);
            await page.waitForTimeout(500);

            await page.keyboard.press('Control+A');
            await page.keyboard.press('Backspace');
            await page.waitForTimeout(300);

            await page.keyboard.type(accountName, { delay: 50 });
            await page.waitForTimeout(500);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(500);

            if (IS_LEDGER_MENU) {
                // 검색 버튼 없음 — 테이블 갱신 대기 후 다운로드
                console.log(`[${accountName}] 테이블 갱신 대기 중 (3초)...`);
                await page.waitForTimeout(3000);
                await handleDownloadAndSave(page, 'button:has-text("엑셀 다운로드")', accountName, rawDataDir, menuName, filePrefix);

            } else {
                // 상세 거래 검색 / 벤포드: 라디오 버튼 선택(옵션) → 검색 → 다운로드
                if (menuName === '상세 거래 검색' && task['표시방식']) {
                    const rbLabel = String(task['표시방식']);
                    try {
                        await page.locator(`label:has-text("${rbLabel}")`).click({ timeout: 5000 });
                        await page.waitForTimeout(500);
                    } catch {
                        console.log(`[경고] 라디오 버튼 '${rbLabel}'을 찾을 수 없습니다.`);
                    }
                }

                const searchBtnSelector = config.selectors.searchButton || 'button:has-text("검색")';
                await page.click(searchBtnSelector);
                await page.waitForTimeout(1000);

                const downloadBtnSelector = config.selectors.excelDownloadBtn || 'button:has-text("결과 다운로드")';
                const targetName = (menuName === '상세 거래 검색')
                    ? accountName
                    : String(task['파일명'] ?? accountName);
                await handleDownloadAndSave(page, downloadBtnSelector, targetName, rawDataDir, menuName, filePrefix);
            }
        }

    } else if (menuName === '매입/매출 이중거래처 분석') {
        console.log(`\n--- [${menuName}] 처리 시작 ---`);
        const task = tasks[0] ?? {};
        await page.click('button:has-text("이중거래처 분석 시작")');
        await page.waitForTimeout(1000);
        const fileName = String(task['파일명'] ?? '이중거래처_결과');
        await handleDownloadAndSave(page, 'button:has-text("결과 다운로드")', fileName, rawDataDir, menuName, filePrefix);

    } else {
        console.log(`[${menuName}] 구현되지 않은 메뉴 형식입니다. 생략합니다.`);
    }
}

// ─── /ai-analysis 엔드포인트 메뉴 핸들러 ────────────────────────────────────
async function handleAiAnalysisMenu(page, menu, config, companyDir, rawDataDir, filePrefix) {
    const { menuName, tasks } = menu;
    console.log(`\n=== [메뉴 진입] ${menuName} (AI 분석) ===`);

    for (const task of tasks) {
        const uploadRelPath = String(task['업로드파일'] ?? task['파일명'] ?? config.uploadFileName ?? '');
        if (!uploadRelPath) {
            console.log(`[${menuName}] 업로드할 파일이 지정되지 않았습니다. 건너뜁니다.`);
            continue;
        }

        const uploadFilePath = path.isAbsolute(uploadRelPath)
            ? uploadRelPath
            : path.join(companyDir, uploadRelPath);

        if (!fs.existsSync(uploadFilePath)) {
            console.log(`[${menuName}] 업로드 파일이 존재하지 않습니다: ${uploadFilePath}`);
            continue;
        }
        console.log(`[${menuName}] 업로드 파일: ${path.basename(uploadFilePath)}`);

        const fileInputSelector = config.selectors.fileUploadInput || 'input[type="file"]';
        const uploadBtnSelector = config.selectors.uploadButton;

        if (uploadBtnSelector) {
            try {
                const [fileChooser] = await Promise.all([
                    page.waitForEvent('filechooser', { timeout: 10000 }),
                    page.click(uploadBtnSelector),
                ]);
                await fileChooser.setFiles(uploadFilePath);
                console.log(`[${menuName}] 파일 선택 완료 (fileChooser 방식).`);
            } catch {
                console.log(`[${menuName}] fileChooser 방식 실패, 직접 주입 방식으로 전환합니다.`);
                await page.waitForSelector(fileInputSelector, { state: 'attached', timeout: 30000 });
                await page.setInputFiles(fileInputSelector, uploadFilePath);
            }
        } else {
            await page.waitForSelector(fileInputSelector, { state: 'attached', timeout: 30000 });
            await page.setInputFiles(fileInputSelector, uploadFilePath);
            console.log(`[${menuName}] 파일 주입 완료 (setInputFiles 방식).`);
        }

        const downloadBtnSelector = config.selectors.excelDownloadBtn || 'button:has-text("결과 다운로드")';
        const outputFileName = String(task['결과파일명'] ?? task['파일명'] ?? `${menuName}_결과`);
        await handleDownloadAndSave(page, downloadBtnSelector, outputFileName, rawDataDir, menuName, filePrefix);
    }
}

// ─── 메인 러너 ────────────────────────────────────────────────────────────────
async function runAudit(config, companyDir) {
    const companyName = config.companyName || path.basename(companyDir);
    console.log(`\n=== ${companyName} 감사 자동화 시작 ===`);

    const isHeadless = config.taskList?.RunMode === 'Debug' ? false : true;
    const clientName = config.taskList?.ClientName ?? config.companyName ?? companyName;
    const targetYear = config.taskList?.TargetYear ?? '';
    const now = new Date();
    const runTimestamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const filePrefix = targetYear ? `${clientName}_${targetYear}_${runTimestamp}_` : `${clientName}_${runTimestamp}_`;

    const persistentProfile = config.persistentProfile
        ?? path.join(__dirname, '..', '.browser_profile');

    const { browser, page } = await initBrowser(isHeadless, persistentProfile);

    try {
        const baseUrl = (config.taskList?.Url ?? config.url ?? '').replace(/\/$/, '');
        if (!baseUrl) {
            throw new Error('접속 URL이 설정되지 않았습니다. config.js 또는 task_list의 Settings 시트를 확인하세요.');
        }
        console.log(`[${companyName}] 접속 URL: ${baseUrl}`);

        // ── 0. 초기 접속 ──────────────────────────────────────────────────────
        await page.goto(baseUrl, { waitUntil: 'networkidle', timeout: 60000 });

        // ── 1. 로그인 (세션이 없을 때만) ─────────────────────────────────────
        const emailSelector = config.selectors.loginId || 'input[type="email"]';
        const loginFormVisible = await page.$(emailSelector)
            .then(el => !!el)
            .catch(() => false);

        if (loginFormVisible && config.credentials?.userId) {
            console.log(`[${companyName}] 로그인 폼 감지. 로그인을 진행합니다...`);
            const pwSelector = config.selectors.loginPassword || 'input[type="password"]';
            const loginBtnSelector = config.selectors.loginButton || 'button:has-text("로그인")';

            await page.waitForSelector(emailSelector, { state: 'visible', timeout: 30000 });
            await page.fill(emailSelector, config.credentials.userId);
            await page.fill(pwSelector, config.credentials.userPassword ?? '');
            await page.click(loginBtnSelector);
            console.log(`[${companyName}] 로그인 완료. 화면 전환 대기 중...`);
            await page.waitForTimeout(2000);
        } else if (!loginFormVisible) {
            console.log(`[${companyName}] 기존 세션 감지. 로그인을 생략합니다.`);
        } else {
            console.log(`[${companyName}] 로그인 정보가 없어 로그인을 생략합니다.`);
        }

        // ── 2. 폴더 준비 ─────────────────────────────────────────────────────
        // raw_data: 웹 앱에 업로드하는 원본 파일 보관
        const rawDataDir = path.join(companyDir, 'raw_data');
        if (!fs.existsSync(rawDataDir)) fs.mkdirSync(rawDataDir, { recursive: true });
        // results: 웹 앱에서 다운로드한 분석 결과물 저장
        const resultsDir = path.join(companyDir, 'results');
        if (!fs.existsSync(resultsDir)) fs.mkdirSync(resultsDir, { recursive: true });

        if (!config.menus?.length) {
            console.log(`[${companyName}] 실행할 메뉴(지시서 시트)가 없습니다. 종료합니다.`);
            return;
        }

        // ── 3. 메뉴 순회 ─────────────────────────────────────────────────────
        let currentEndpoint = null;
        let analysisUploadDone = false; // /analysis 파일 업로드는 한 번만

        for (const menu of config.menus) {
            const menuName = menu.menuName;
            const endpoint = getMenuEndpoint(menuName, config);
            const targetUrl = `${baseUrl}${endpoint}`;

            // 엔드포인트가 바뀔 때만 페이지 이동
            if (currentEndpoint !== endpoint) {
                console.log(`\n[라우팅] ${menuName} → ${targetUrl}`);
                await page.goto(targetUrl, { waitUntil: 'networkidle', timeout: 60000 });
                currentEndpoint = endpoint;
                analysisUploadDone = false; // 페이지 이동 시 업로드 상태 초기화
                await page.waitForTimeout(1000);
            }

            if (endpoint === '/ai-analysis') {
                // 분개장 등 AI 분석 메뉴 (업로드는 핸들러 내부에서 task별로 처리)
                await handleAiAnalysisMenu(page, menu, config, companyDir, resultsDir, filePrefix);

            } else {
                // /analysis 메뉴
                // 1) 계정별원장 파일 업로드 (최초 1회)
                if (!analysisUploadDone && config.uploadFileName) {
                    await uploadGeneralLedgerIfNeeded(page, config, companyDir);
                    analysisUploadDone = true;
                }

                // 2) 분석 메뉴 카드 클릭
                // 시트명과 UI 카드 텍스트가 다를 수 있으므로 매핑 테이블 우선 조회
                const uiLabel = getMenuUiLabel(menuName, config);
                console.log(`\n=== [메뉴 진입] ${menuName}${uiLabel !== menuName ? ` → UI: "${uiLabel}"` : ''} ===`);

                // 정확한 텍스트 매칭 우선, 없으면 부분 포함 매칭으로 폴백
                let menuHandle = await page.$(`text="${uiLabel}"`).catch(() => null);
                if (!menuHandle) {
                    menuHandle = await page.$(`h2:has-text("${uiLabel}"), h3:has-text("${uiLabel}"), span:has-text("${uiLabel}"), div:has-text("${uiLabel}")`).catch(() => null);
                }

                if (menuHandle) {
                    await menuHandle.evaluate(node => {
                        node.removeAttribute('target');
                        node.closest('a')?.removeAttribute('target');
                    });
                    await menuHandle.click();
                    await page.waitForTimeout(2000);
                } else {
                    console.log(`[경고] UI에서 "${uiLabel}" 카드/버튼을 찾지 못했습니다. 현재 화면에서 바로 처리합니다.`);
                }

                // 3) 계정별 데이터 추출
                await handleAnalysisMenu(page, menu, config, resultsDir, filePrefix);
            }
        }

        console.log(`\n=== ${companyName} 자동화 완료 ===`);
        console.log(`최종 결과물이 ${companyName}/results 폴더에 저장되었습니다.`);

    } catch (error) {
        console.error(`[${companyName}] 실행 중 오류 발생:`, error.message);
        try {
            const screenshotPath = path.join(companyDir, 'error.png');
            await page.screenshot({ path: screenshotPath, fullPage: true });
            console.log(`[${companyName}] 에러 스크린샷 저장: ${screenshotPath}`);
        } catch {
            console.error(`[${companyName}] 스크린샷 저장 실패.`);
        }
    } finally {
        await browser.close();
    }
}

module.exports = runAudit;
