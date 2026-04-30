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

// host1_xxx / HOST1_xxx → /analysis, host2_xxx / HOST2_xxx → /ai-analysis
function getBaseMenuName(menuName) {
    return menuName.replace(/^host_?\d+_?/i, '');
}

function getMenuEndpoint(menuName, config) {
    if (/^host_?2_?/i.test(menuName)) return '/ai-analysis';
    const base = getBaseMenuName(menuName);
    return config.menuEndpoints?.[menuName]
        ?? config.menuEndpoints?.[base]
        ?? DEFAULT_ENDPOINT_MAP[menuName]
        ?? DEFAULT_ENDPOINT_MAP[base]
        ?? '/analysis';
}

// ─── 시트명 → UI 카드/버튼 텍스트 매핑 ──────────────────────────────────────
// task_list.xlsx의 시트명(menuName)과 웹 앱의 실제 카드 텍스트가 다를 때 사용.
// config.menuLabels 로 회사별 커스텀 오버라이드 가능.
const DEFAULT_MENU_LABEL_MAP = {
    '총계정원장':             '총계정원장 조회',
    '상세거래검색':           '상세 거래 검색',
    '상세검색_시나리오':      '상세 거래 검색',   // 시나리오 시트 → 동일 UI 카드 진입
    '이중거래처분석':         '매입/매출 이중거래처 분석',
    '벤포드':                 '벤포드 법칙 분석',
    '벤포드법칙분석':         '벤포드 법칙 분석',
    '벤포드법칙':             '벤포드 법칙 분석',
    '계정연관거래처':         '계정 연관 거래처 분석',
    '외상매출매입상계':       '외상매출/매입 상계 거래처 분석',
    '추정손익':               '추정 손익 분석',
    '매출관리비추이':         '매출/관리비 월별 추이 분석',
    '전기비교':               '전기 데이터 비교 분석',
    '감사샘플링':             '감사 샘플링',
    '금감원위험분석':         '금감원 지적사례 기반 위험 분석',
    '재무제표증감':           '재무제표 증감 분석',
    'HOST2_월별트렌드분석':   '월별 트렌드 분석',
    'HOST2_월별트렌드':       '월별 트렌드 분석',
};

function getMenuUiLabel(menuName, config) {
    const base = getBaseMenuName(menuName);
    return config.menuLabels?.[menuName]
        ?? config.menuLabels?.[base]
        ?? DEFAULT_MENU_LABEL_MAP[menuName]
        ?? DEFAULT_MENU_LABEL_MAP[base]
        ?? base;
}

// ─── 전기(전년도) 계정별원장 파일 탐색 ──────────────────────────────────────────
// raw_data 폴더에서 파일명에 '전기'가 명시적으로 포함된 xlsx 파일만 반환.
function findPrevYearLedgerFile(rawDataDir) {
    if (!fs.existsSync(rawDataDir)) return null;

    const found = fs.readdirSync(rawDataDir).find(f => {
        const ext = path.extname(f).toLowerCase();
        return f.includes('전기') && (ext === '.xlsx' || ext === '.xls');
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

// ─── 날짜 포맷 헬퍼 ──────────────────────────────────────────────────────────
// ExcelJS는 날짜 셀을 JS Date 객체로 반환함. 문자열(YYYY-MM-DD 등)도 허용.
function formatExcelDate(val) {
    if (!val) return '';
    if (val instanceof Date) {
        const y = val.getFullYear();
        const m = String(val.getMonth() + 1).padStart(2, '0');
        const d = String(val.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }
    return String(val).trim();
}

// ─── 텍스트 입력창 채우기 헬퍼 ────────────────────────────────────────────────
// labelText(한글 레이블)와 연결된 input을 여러 전략으로 탐색 후 값 입력.
// 성공 시 true 반환.
async function tryFillInput(page, labelText, value) {
    // 전략 1: aria-label / htmlFor 연결 (Playwright getByLabel)
    try {
        const loc = page.getByLabel(new RegExp(labelText, 'i'));
        if (await loc.count() > 0) {
            await loc.first().clear();
            await loc.first().fill(value);
            return true;
        }
    } catch { /* 다음 전략으로 */ }

    // 전략 2: placeholder 포함
    try {
        const loc = page.locator(`input[placeholder*="${labelText}"]`);
        if (await loc.count() > 0) {
            await loc.first().clear();
            await loc.first().fill(value);
            return true;
        }
    } catch { /* 다음 전략으로 */ }

    // 전략 3: label 요소가 input을 감싸거나 인접한 경우
    try {
        const loc = page.locator(
            `label:has-text("${labelText}") input, ` +
            `label:has-text("${labelText}") + input, ` +
            `label:has-text("${labelText}") ~ input`
        );
        if (await loc.count() > 0) {
            await loc.first().clear();
            await loc.first().fill(value);
            return true;
        }
    } catch { /* 실패 */ }

    return false;
}

// ─── 날짜 입력 헬퍼 ───────────────────────────────────────────────────────────
// type="date" input은 fill('YYYY-MM-DD') 로 직접 처리.
// labelKeyword: '시작', '종료' 등 레이블에 포함된 키워드
async function fillDateInput(page, labelKeyword, dateStr) {
    if (!dateStr) return;
    try {
        const byLabel = page.getByLabel(new RegExp(labelKeyword, 'i'));
        if (await byLabel.count() > 0) {
            await byLabel.first().fill(dateStr);
            await byLabel.first().press('Tab'); // 변경 이벤트 트리거
            return;
        }
    } catch { /* 다음 전략 */ }
    try {
        const loc = page.locator(`input[placeholder*="${labelKeyword}"], input[type="date"]`).first();
        if (await loc.count() > 0) {
            await loc.fill(dateStr);
            await loc.press('Tab');
        }
    } catch (e) {
        console.log(`[경고] '${labelKeyword}' 날짜 입력 실패: ${e.message}`);
    }
}

// ─── 라디오 버튼 클릭 헬퍼 ───────────────────────────────────────────────────
// 엑셀의 텍스트와 화면의 라디오 버튼 레이블 텍스트를 직접 매칭하여 클릭.
async function clickRadioByLabel(page, labelText, groupHint) {
    if (!labelText) return;
    const text = String(labelText).trim();
    const exactRe = new RegExp(`^${text}$`);

    // label → button(정확) → button(포함) → tab → role=radio 순으로 시도
    const candidates = [
        page.locator('label').filter({ hasText: exactRe }),
        page.locator(`label:has-text("${text}")`),
        page.locator('button').filter({ hasText: exactRe }),
        page.locator(`button:has-text("${text}")`),
        page.locator(`[role="tab"]:has-text("${text}")`),
        page.locator(`[role="radio"]:has-text("${text}")`),
    ];

    for (const loc of candidates) {
        try {
            if (await loc.count().catch(() => 0) === 0) continue;
            await loc.first().click({ timeout: 3000 });
            await page.waitForTimeout(300);
            console.log(`  ✓ '${text}' 선택`);
            return;
        } catch { /* 다음 셀렉터 */ }
    }
    console.log(`[경고] '${groupHint ?? ''}' 항목 '${text}' 클릭 실패 — 건너뜁니다.`);
}

// ─── 다운로드 → workbook 시트 추가 헬퍼 ─────────────────────────────────────
// 결과 다운로드 후 sheetName 으로 workbook에 시트를 추가. 파일은 저장하지 않음.
async function downloadAndAddSheet(page, downloadBtnSelector, sheetName, workbook, menuName) {
    console.log(`[${menuName}] '${sheetName}' 결과 다운로드 대기 중...`);
    await page.waitForSelector(downloadBtnSelector, { state: 'visible', timeout: 30000 });

    const downloadPromise = page.waitForEvent('download');
    await page.click(downloadBtnSelector);
    const download = await downloadPromise;
    const downloadPath = await download.path();
    console.log(`[${menuName}] 다운로드 캡처 완료.`);

    const safeSheetName = sheetName.substring(0, 31).replace(/[\\/?*[\]:]/g, '_');
    const srcBook = new ExcelJS.Workbook();
    await srcBook.xlsx.readFile(downloadPath);
    const srcSheet = srcBook.worksheets[0];

    if (workbook.getWorksheet(safeSheetName)) workbook.removeWorksheet(safeSheetName);
    const destSheet = workbook.addWorksheet(safeSheetName);
    srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const destRow = destSheet.getRow(rowNumber);
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            destRow.getCell(colNumber).value = cell.value;
        });
        destRow.commit();
    });
    console.log(`[${menuName}] '${safeSheetName}' 시트 추가 완료.`);
}

// ─── 상세검색_시나리오 전용 핸들러 ───────────────────────────────────────────
// 동일 '작업명' 행들을 하나의 엑셀 파일로 묶고, 계정과목명을 시트명으로 사용.
// 각 계정 처리 후 '뒤로가기'로 검색 화면으로 복귀하여 다음 계정을 이어서 처리.
async function handleDetailSearchScenario(page, menu, config, resultsDir, filePrefix) {
    const { menuName, tasks } = menu;

    // ── 작업명 기준 그룹화 (순서 유지) ──────────────────────────────────────
    const taskGroups = new Map();
    for (const task of tasks) {
        const taskName    = String(task['작업명']   ?? '').trim();
        const accountName = String(task['계정과목'] ?? '').trim();
        if (!taskName && !accountName) continue;
        const key = taskName || accountName;
        if (!taskGroups.has(key)) taskGroups.set(key, []);
        taskGroups.get(key).push(task);
    }

    const allGroups = [...taskGroups.entries()];

    for (let gi = 0; gi < allGroups.length; gi++) {
        const [taskName, groupTasks] = allGroups[gi];
        console.log(`\n=== [작업그룹: ${taskName}] ${groupTasks.length}개 계정 처리 시작 ===`);

        const groupBook    = new ExcelJS.Workbook();
        const safeFileName = taskName.substring(0, 50).replace(/[\\/?*[\]:]/g, '_');
        const groupFilePath = path.join(resultsDir, `${filePrefix}${safeFileName}.xlsx`);

        // 상세 거래 검색 카드 UI 레이블 (뒤로가기 후 재진입에 사용)
        const cardUiLabel = getMenuUiLabel(menuName, config);
        const comboSel    = config.selectors.accountCombobox || 'button[role="combobox"]';

        for (let ti = 0; ti < groupTasks.length; ti++) {
            const task        = groupTasks[ti];
            const accountName = String(task['계정과목'] ?? '').trim();
            const vendorName  = String(task['거래처명'] ?? '').trim();
            const description = String(task['적요']     ?? '').trim();
            const amountType  = String(task['금액유형'] ?? task['금액 유형'] ?? '').trim();
            const displayType = String(task['표시방식'] ?? task['표시 방식'] ?? '').trim();
            const startDateRaw = task['시작일'] ?? task['시작일자'] ?? task['기간시작'] ?? null;
            const endDateRaw   = task['종료일'] ?? task['종료일자'] ?? task['기간종료'] ?? null;

            if (!accountName) {
                console.log(`[${taskName}] 계정과목 없음 — 건너뜁니다.`);
                continue;
            }
            console.log(`\n--- [${taskName} / ${accountName}] 처리 시작 ---`);

            try {
                // 1. 콤보박스가 보일 때까지 대기 후 계정과목 입력
                // (초기화는 이전 반복 종료 시점에 수행되므로 여기서는 생략)
                await page.waitForSelector(comboSel, { state: 'visible', timeout: 10000 });
                await page.waitForTimeout(500);
                await page.click(comboSel);
                await page.waitForTimeout(300);
                await page.keyboard.press('Control+A');
                await page.keyboard.press('Backspace');
                await page.keyboard.type(accountName, { delay: 50 });
                await page.waitForTimeout(500);
                await page.keyboard.press('Enter');
                await page.waitForTimeout(400);
                console.log(`  계정과목: ${accountName}`);

                // 3. 거래처명 — 두 번째 combobox: 항상 먼저 지우고, 값 있으면 입력
                try {
                    const vendorCombo = page.locator(comboSel).nth(1);
                    await vendorCombo.click();
                    await page.waitForTimeout(200);
                    await page.keyboard.press('Control+A');
                    await page.keyboard.press('Backspace');
                    if (vendorName) {
                        await page.keyboard.type(vendorName, { delay: 50 });
                        await page.waitForTimeout(400);
                        await page.keyboard.press('Enter');
                        await page.waitForTimeout(300);
                        console.log(`  거래처명: ${vendorName}`);
                    } else {
                        await page.keyboard.press('Escape');
                        await page.waitForTimeout(200);
                    }
                } catch {
                    console.log(`[경고] 거래처명 combobox를 찾지 못했습니다.`);
                }

                // 4. 적요 — 세 번째 combobox: 항상 먼저 지우고, 값 있으면 입력
                try {
                    const descCombo = page.locator(comboSel).nth(2);
                    await descCombo.click();
                    await page.waitForTimeout(200);
                    await page.keyboard.press('Control+A');
                    await page.keyboard.press('Backspace');
                    if (description) {
                        await page.keyboard.type(description, { delay: 50 });
                        await page.waitForTimeout(400);
                        await page.keyboard.press('Enter');
                        await page.waitForTimeout(300);
                        console.log(`  적요: ${description}`);
                    } else {
                        await page.keyboard.press('Escape');
                        await page.waitForTimeout(200);
                    }
                } catch {
                    console.log(`[경고] 적요 combobox를 찾지 못했습니다.`);
                }

                // 5. 날짜
                if (startDateRaw) { const d = formatExcelDate(startDateRaw); await fillDateInput(page, '시작', d); console.log(`  시작일: ${d}`); }
                if (endDateRaw)   { const d = formatExcelDate(endDateRaw);   await fillDateInput(page, '종료', d); console.log(`  종료일: ${d}`); }

                // 6. 금액 유형 / 표시 방식 라디오 (행마다 개별 적용)
                if (amountType)  await clickRadioByLabel(page, amountType,  '금액 유형');
                if (displayType) await clickRadioByLabel(page, displayType, '표시 방식');

                // 7. 검색
                const searchSel = config.selectors.searchButton || 'button:has-text("검색")';
                await page.waitForSelector(searchSel, { state: 'visible', timeout: 10000 });
                await page.click(searchSel);
                await page.waitForTimeout(1500);

                // 8. 다운로드 → 그룹 workbook에 시트 추가
                const downloadSel = config.selectors.excelDownloadBtn || 'button:has-text("결과 다운로드")';
                const dlVisible = await page.locator(downloadSel).waitFor({ state: 'visible', timeout: 8000 }).then(() => true).catch(() => false);
                if (!dlVisible) {
                    console.log(`  [안내] '${accountName}' 검색 결과 없음 — 다운로드를 건너뜁니다.`);
                } else {
                    await downloadAndAddSheet(page, downloadSel, accountName, groupBook, menuName);
                }

            } catch (e) {
                console.log(`[경고] [${taskName} / ${accountName}] 처리 중 오류 (다음 계정으로 진행): ${e.message}`);
            }

            // 9. 다음 태스크가 있으면: [초기화] 버튼 클릭 → 폼 안정화 → 다음 계정 입력 준비
            const isLastOverall = gi === allGroups.length - 1 && ti === groupTasks.length - 1;
            if (!isLastOverall) {
                const resetSel = config.selectors.resetButton;
                if (resetSel) {
                    try {
                        await page.waitForSelector(resetSel, { state: 'visible', timeout: 5000 });
                        await page.click(resetSel);
                        await page.waitForTimeout(1000); // 폼 초기화 안정화 대기
                        await page.waitForSelector(comboSel, { state: 'visible', timeout: 10000 });
                        console.log(`  → [초기화] 완료, 다음 계정 입력 준비`);
                    } catch (e) {
                        console.log(`[경고] 초기화 버튼 클릭 실패 (다음 계정으로 진행): ${e.message}`);
                    }
                } else {
                    console.log(`[경고] config.selectors.resetButton 이 설정되지 않았습니다. 초기화를 건너뜁니다.`);
                }
            }
        }

        // 10. 그룹 파일 저장 (OneDrive EBUSY 재시도 포함)
        for (let attempt = 1; attempt <= 5; attempt++) {
            try {
                await groupBook.xlsx.writeFile(groupFilePath);
                console.log(`[${taskName}] 그룹 파일 저장 완료: ${path.basename(groupFilePath)}`);
                break;
            } catch (e) {
                if (e.code === 'EBUSY' && attempt < 5) {
                    console.log(`[${taskName}] 파일 잠금 감지, ${attempt}초 후 재시도...`);
                    await new Promise(r => setTimeout(r, attempt * 1000));
                } else { throw e; }
            }
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
        // OneDrive 동기화로 인한 파일 잠금(EBUSY) 대비 재시도
        let copied = false;
        for (let attempt = 1; attempt <= 5; attempt++) {
            try {
                fs.copyFileSync(downloadPath, finalPath);
                copied = true;
                break;
            } catch (e) {
                if (e.code === 'EBUSY' && attempt < 5) {
                    console.log(`[${menuName}] 파일 잠금 감지, ${attempt}초 후 재시도... (${attempt}/5)`);
                    await new Promise(r => setTimeout(r, attempt * 1000));
                } else {
                    throw e;
                }
            }
        }
        if (copied) console.log(`[${menuName}] 개별 파일 저장 완료: ${path.basename(finalPath)}`);
    }
}

// ─── /analysis 엔드포인트 메뉴 핸들러 ────────────────────────────────────────
// 시트명("총계정원장", "총계정원장조회" 등)을 모두 처리.
async function handleAnalysisMenu(page, menu, config, rawDataDir, filePrefix) {
    const { menuName, tasks } = menu;
    const base = getBaseMenuName(menuName);

    // ── 상세검색_시나리오 시트: 전용 핸들러로 위임 ───────────────────────────
    if (base === '상세검색_시나리오') {
        return handleDetailSearchScenario(page, menu, config, rawDataDir, filePrefix);
    }

    // "총계정원장" 계열: 계정 콤보박스 → 대기 → 다운로드
    const IS_LEDGER_MENU = ['총계정원장', '총계정원장조회'].includes(base);
    // "벤포드 법칙 분석": 계정 선택 + 금액기준열 선택 + "분석 시작" 버튼
    const IS_BENFORD_MENU = ['벤포드법칙분석', '벤포드법칙', '벤포드', '벤포드 법칙 분석'].includes(base);
    // "상세 거래 검색": 계정 선택 + 검색 버튼 → 다운로드
    const IS_SEARCH_MENU = ['상세 거래 검색'].includes(base);

    if (IS_BENFORD_MENU) {
        for (const task of tasks) {
            const taskKeys = Object.keys(task);
            if (taskKeys.length === 0) continue;

            const accountName = String(task['계정과목'] ?? task[taskKeys[0]] ?? '').trim();
            if (!accountName) continue;

            console.log(`\n--- [벤포드 / ${accountName}] 처리 시작 ---`);

            const comboboxSelector = config.selectors.accountCombobox || 'button[role="combobox"]';

            // 1) 계정과목 선택 (첫 번째 combobox)
            const combos = page.locator(comboboxSelector);
            const comboCount = await combos.count().catch(() => 0);
            const accountCombo = comboCount > 0 ? combos.first() : null;
            if (accountCombo) {
                await accountCombo.click();
                await page.waitForTimeout(500);
                await page.keyboard.press('Control+A');
                await page.keyboard.press('Backspace');
                await page.keyboard.type(accountName, { delay: 50 });
                await page.waitForTimeout(500);
                await page.keyboard.press('Enter');
                await page.waitForTimeout(500);
                console.log(`  ✓ 계정과목 '${accountName}' 선택`);
            }

            // 2) 금액 기준열 선택 (task에 '금액기준열' 또는 '기준열' 컬럼이 있으면 적용)
            const amountCol = String(task['금액기준열'] ?? task['기준열'] ?? '').trim();
            if (amountCol && comboCount >= 2) {
                const colCombo = combos.nth(1);
                await colCombo.click();
                await page.waitForTimeout(400);
                try {
                    await page.locator(`[role="option"]:has-text("${amountCol}")`).first().click({ timeout: 3000 });
                    console.log(`  ✓ 금액 기준열 '${amountCol}' 선택`);
                } catch {
                    await page.keyboard.press('Escape');
                    console.log(`  [경고] 금액 기준열 '${amountCol}'을 찾지 못했습니다. 기본값 유지.`);
                }
                await page.waitForTimeout(400);
            }

            // 3) 분석 시작 클릭
            await page.click('button:has-text("분석 시작")');
            await page.waitForTimeout(2000);
            console.log(`  ✓ 분석 시작 클릭`);

            // 4) 결과 다운로드 (벤포드 결과 섹션의 "엑셀 다운로드" 버튼)
            const targetName = String(task['파일명'] ?? accountName);
            const dlBtn = config.selectors.benfordDownloadBtn || 'button:has-text("엑셀 다운로드")';
            await handleDownloadAndSave(page, dlBtn, targetName, rawDataDir, menuName, filePrefix);

            // 5) 뒤로가기로 복귀 (다음 계정 처리를 위해)
            if (tasks.indexOf(task) < tasks.length - 1) {
                try {
                    await page.click('button:has-text("뒤로가기"), a:has-text("뒤로가기")', { timeout: 3000 });
                    await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
                    await page.waitForTimeout(1000);
                } catch { /* 뒤로가기 없으면 무시 */ }
            }
        }

    } else if (IS_LEDGER_MENU || IS_SEARCH_MENU) {
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
                if (base === '상세 거래 검색' && task['표시방식']) {
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
                const targetName = (base === '상세 거래 검색')
                    ? accountName
                    : String(task['파일명'] ?? accountName);
                await handleDownloadAndSave(page, downloadBtnSelector, targetName, rawDataDir, menuName, filePrefix);
            }
        }

    } else if (['이중거래처분석', '매입/매출 이중거래처 분석'].includes(base)) {
        console.log(`\n--- [${menuName}] 처리 시작 ---`);
        const task = tasks[0] ?? {};
        await page.click('button:has-text("이중거래처 분석 시작")');
        await page.waitForTimeout(1000);
        const fileName = String(task['파일명'] ?? base);
        await handleDownloadAndSave(page, 'button:has-text("결과 다운로드")', fileName, rawDataDir, menuName, filePrefix);

    } else if (['외상매출매입상계', '외상매출/매입 상계 거래처 분석'].includes(base)) {
        console.log(`\n--- [${menuName}] 처리 시작 ---`);
        const task = tasks[0] ?? {};
        await page.click('button:has-text("상계 거래처 분석 시작")');
        await page.waitForTimeout(1000);
        const fileName = String(task['파일명'] ?? base);
        await handleDownloadAndSave(page, 'button:has-text("결과 다운로드")', fileName, rawDataDir, menuName, filePrefix);

    } else if (['전기비교', '전기 데이터 비교 분석'].includes(base)) {
        const comboboxSelector = config.selectors.accountCombobox || 'button[role="combobox"]';
        for (const task of tasks) {
            const accountName = String(task['분석할 계정과목'] ?? task['계정과목'] ?? '').trim();
            const amountType  = String(task['금액 기준열']    ?? task['금액유형']   ?? '').trim();
            if (!accountName) continue;

            console.log(`\n--- [전기비교 / ${accountName}] 처리 시작 ---`);

            // 1. 계정명 combobox 선택 (자동 분석 트리거)
            await page.waitForSelector(comboboxSelector, { state: 'visible', timeout: 10000 });
            await page.click(comboboxSelector);
            await page.waitForTimeout(300);
            await page.keyboard.press('Control+A');
            await page.keyboard.press('Backspace');
            await page.keyboard.type(accountName, { delay: 50 });
            await page.waitForTimeout(500);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1500);
            console.log(`  ✓ 계정 '${accountName}' 선택`);

            // 2. 금액 유형 라디오 (차변만 / 대변만 / 차변+대변 모두)
            if (amountType) {
                await clickRadioByLabel(page, amountType, '금액 유형');
                await page.waitForTimeout(1000);
            }

            // 3. 비교표 다운로드
            const targetName = String(task['파일명'] ?? `전기비교_${accountName}`);
            await handleDownloadAndSave(page, 'button:has-text("비교표 다운로드")', targetName, rawDataDir, menuName, filePrefix);

            // 4. 다음 계정 처리 위해 초기화
            if (tasks.indexOf(task) < tasks.length - 1) {
                try {
                    await page.click('button:has-text("초기화"), a:has-text("초기화")', { timeout: 3000 });
                    await page.waitForTimeout(800);
                } catch { /* 초기화 버튼 없으면 무시 */ }
            }
        }

    } else {
        console.log(`[${menuName}] 구현되지 않은 메뉴 형식입니다. 생략합니다.`);
    }
}

// ─── 월별 이상치 감지 ─────────────────────────────────────────────────────────
// monthlyData: [{ month: 'YYYY-MM', debit: number, credit: number }, ...]
// threshold: 0.3 → 평균 대비 30% 초과 시 이상치
function detectMonthlyAnomalies(monthlyData, threshold = 0.3) {
    const avg = arr => arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0;

    const debitAmounts  = monthlyData.map(m => m.debit).filter(v => v > 0);
    const creditAmounts = monthlyData.map(m => m.credit).filter(v => v > 0);
    const debitAvg  = avg(debitAmounts);
    const creditAvg = avg(creditAmounts);

    console.log(`[이상치감지] 차변 월평균: ${Math.round(debitAvg).toLocaleString()}, 대변 월평균: ${Math.round(creditAvg).toLocaleString()}`);

    const anomalies = [];
    for (const m of monthlyData) {
        if (debitAvg > 0 && m.debit > debitAvg * (1 + threshold)) {
            const pct = ((m.debit / debitAvg - 1) * 100).toFixed(1);
            console.log(`  ★ 차변 급증 — ${m.month}: ${m.debit.toLocaleString()} (평균 대비 +${pct}%)`);
            anomalies.push({ month: m.month, type: '차변', amount: m.debit, avg: debitAvg });
        }
        if (debitAvg > 0 && m.debit > 0 && m.debit < debitAvg * (1 - threshold)) {
            const pct = ((1 - m.debit / debitAvg) * 100).toFixed(1);
            console.log(`  ★ 차변 급감 — ${m.month}: ${m.debit.toLocaleString()} (평균 대비 -${pct}%)`);
            anomalies.push({ month: m.month, type: '차변', amount: m.debit, avg: debitAvg });
        }
        if (creditAvg > 0 && m.credit > creditAvg * (1 + threshold)) {
            const pct = ((m.credit / creditAvg - 1) * 100).toFixed(1);
            console.log(`  ★ 대변 급증 — ${m.month}: ${m.credit.toLocaleString()} (평균 대비 +${pct}%)`);
            anomalies.push({ month: m.month, type: '대변', amount: m.credit, avg: creditAvg });
        }
        if (creditAvg > 0 && m.credit > 0 && m.credit < creditAvg * (1 - threshold)) {
            const pct = ((1 - m.credit / creditAvg) * 100).toFixed(1);
            console.log(`  ★ 대변 급감 — ${m.month}: ${m.credit.toLocaleString()} (평균 대비 -${pct}%)`);
            anomalies.push({ month: m.month, type: '대변', amount: m.credit, avg: creditAvg });
        }
    }
    return anomalies;
}

// ─── 월별 금액 데이터 추출 (다중 전략) ────────────────────────────────────────
async function extractMonthlyAmountsFromPage(page, menuName) {
    // 전략 1: 요약 테이블 파싱 (월 | 차변금액 | 대변금액 형태)
    try {
        const rows = await page.$$eval('table tr', rows =>
            rows.map(row => {
                const cells = [...row.querySelectorAll('td, th')].map(c => c.innerText?.trim() ?? '');
                const monthMatch = cells[0]?.match(/\d{4}-\d{2}/);
                if (!monthMatch) return null;
                const parseNum = t => Number((t ?? '').replace(/[^0-9.-]/g, '')) || 0;
                return { month: monthMatch[0], debit: parseNum(cells[1]), credit: parseNum(cells[2]) };
            }).filter(Boolean)
        );
        if (rows.length > 0) {
            console.log(`[${menuName}] 요약 테이블 파싱: ${rows.length}개월 추출`);
            return rows;
        }
    } catch { /* 다음 전략 */ }

    // 전략 2: Chart.js 인스턴스 데이터 (v2/v3 모두 시도)
    try {
        const chartData = await page.evaluate(() => {
            const instances = window.Chart?.instances
                ? Object.values(window.Chart.instances)
                : [];
            for (const chart of instances) {
                const labels   = chart.data?.labels ?? [];
                const datasets = chart.data?.datasets ?? [];
                if (!labels.length) continue;
                const debitDs  = datasets.find(d => /차변|debit/i.test(d.label ?? ''));
                const creditDs = datasets.find(d => /대변|credit/i.test(d.label ?? ''));
                if (!debitDs && !creditDs) continue;
                return labels.map((label, i) => ({
                    month:  String(label),
                    debit:  Number(debitDs?.data?.[i]  ?? 0),
                    credit: Number(creditDs?.data?.[i] ?? 0),
                }));
            }
            return null;
        });
        if (chartData?.length > 0) {
            console.log(`[${menuName}] Chart.js 인스턴스 파싱: ${chartData.length}개월 추출`);
            return chartData;
        }
    } catch { /* 다음 전략 */ }

    // 전략 3: data-* 속성 또는 클래스 기반 DOM 파싱
    try {
        const items = await page.$$eval(
            '[data-month], [class*="month-item"], [class*="trend-row"], [class*="monthly"]',
            els => els.map(el => {
                const txt       = el.dataset.month ?? el.querySelector('[class*="month"]')?.innerText ?? '';
                const monthMatch = txt.match(/\d{4}-\d{2}/);
                if (!monthMatch) return null;
                const nums = [...el.querySelectorAll('[class*="amount"], [class*="debit"], [class*="credit"], td')]
                    .map(e => Number((e.innerText ?? '').replace(/[^0-9.-]/g, '')) || 0);
                return { month: monthMatch[0], debit: nums[0] ?? 0, credit: nums[1] ?? 0 };
            }).filter(Boolean)
        );
        if (items.length > 0) {
            console.log(`[${menuName}] DOM 속성 파싱: ${items.length}개월 추출`);
            return items;
        }
    } catch { /* 실패 */ }

    console.log(`[${menuName}] 월별 금액 자동 추출 실패 — 빈 배열 반환`);
    return [];
}

// ─── TOP 10 섹션 드롭다운 선택 헬퍼 ──────────────────────────────────────────
// labelText: '월' | '금액 기준' | '상위'
// optionValue: 실제 선택할 텍스트 값 (예: '2025-01', '차변', 'Top 10')
async function selectTop10FilterDropdown(page, labelText, optionValue, menuName) {
    console.log(`[${menuName}] TOP10 필터 — '${labelText}' → '${optionValue}'`);

    // 전략 1: label 인접 native <select>
    const labelSelectors = [
        `label:has-text("${labelText}") + select`,
        `label:has-text("${labelText}") ~ select`,
        `div:has(> label:has-text("${labelText}")) select`,
        `th:has-text("${labelText}") + th select`,
        `span:has-text("${labelText}") + select`,
        `span:has-text("${labelText}") ~ select`,
    ];
    for (const sel of labelSelectors) {
        try {
            const el = page.locator(sel).first();
            if (await el.count() > 0) {
                await el.selectOption({ label: optionValue });
                await page.waitForTimeout(1200);
                console.log(`  ✓ '${labelText}' 네이티브 select 설정 완료`);
                return;
            }
        } catch { /* 다음 셀렉터 */ }
    }

    // 전략 2: 커스텀 드롭다운 (버튼/div 클릭 → listbox 옵션 클릭)
    const triggerSelectors = [
        `label:has-text("${labelText}") + button`,
        `label:has-text("${labelText}") ~ button`,
        `div:has(> label:has-text("${labelText}")) button`,
        `[aria-label*="${labelText}"]`,
        `button[aria-haspopup="listbox"]:near(label:has-text("${labelText}"))`,
    ];
    for (const sel of triggerSelectors) {
        try {
            const btn = page.locator(sel).first();
            if (await btn.count() > 0) {
                await btn.click();
                await page.waitForTimeout(500);
                await page.click(
                    `[role="listbox"] [role="option"]:has-text("${optionValue}"), ` +
                    `ul[role="listbox"] li:has-text("${optionValue}"), ` +
                    `div[role="option"]:has-text("${optionValue}")`
                );
                await page.waitForTimeout(1200);
                console.log(`  ✓ '${labelText}' 커스텀 드롭다운 설정 완료`);
                return;
            }
        } catch { /* 다음 셀렉터 */ }
    }

    console.log(`[경고] [${menuName}] '${labelText}' 드롭다운을 찾지 못했습니다.`);
}

// ─── 월별 트렌드 이상치 핸들러 ────────────────────────────────────────────────
// HOST2_ 계열 '월별트렌드분석' 시나리오 전용.
// 업로드 완료 후 '월별 트렌드 분석' 카드 진입 → 이상치 감지 → TOP 10 조건부 다운로드.
async function handleMonthlyTrendAnalysis(page, menu, config, companyDir, resultsDir, filePrefix) {
    const { menuName } = menu;
    const companyName = config.companyName ?? path.basename(companyDir);

    // 1. 분석 카드 목록에서 '월별 트렌드 분석' 카드 클릭
    console.log(`[${menuName}] '월별 트렌드 분석' 카드 클릭 시도...`);
    try {
        const cardLoc = page.locator(
            'text=월별 트렌드 분석, ' +
            'text=월별트렌드분석, ' +
            'text=월별 트랜드 분석'
        ).first();
        await cardLoc.waitFor({ state: 'visible', timeout: 15000 });
        await cardLoc.click();
        await page.waitForTimeout(2000);
    } catch (e) {
        console.log(`[경고] [${menuName}] '월별 트렌드 분석' 카드 클릭 실패: ${e.message}`);
    }

    // 2. 페이지 안정화 대기
    await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
    await page.waitForTimeout(1000);

    // 3. 월별 금액 데이터 추출
    const monthlyData = await extractMonthlyAmountsFromPage(page, menuName);
    if (monthlyData.length === 0) {
        console.log(`[${menuName}] 월별 데이터를 읽지 못했습니다. 처리 종료.`);
        return;
    }

    // 4. 이상치 감지 (평균 대비 30% 초과)
    const anomalies = detectMonthlyAnomalies(monthlyData, 0.3);
    if (anomalies.length === 0) {
        console.log(`[${menuName}] 이상치 없음 (기준: 평균 +30%). 처리 종료.`);
        return;
    }
    console.log(`\n[${menuName}] === 이상치 총 ${anomalies.length}건 → TOP 10 추출 시작 ===\n`);

    // 5. '월별 거래처 Top 10' 섹션으로 스크롤
    try {
        await page.locator(
            'text=월별 거래처 Top 10, text=월별 거래처 TOP 10'
        ).first().scrollIntoViewIfNeeded();
        await page.waitForTimeout(1000);
    } catch { /* 스크롤 실패 무시 */ }

    // 6. 이상치별 필터 조작 → 다운로드
    for (const anomaly of anomalies) {
        console.log(`\n--- [${anomaly.month} / ${anomaly.type}] TOP 10 추출 ---`);

        // 월 드롭다운 선택
        await selectTop10FilterDropdown(page, '월', anomaly.month, menuName);

        // 차/대변 드롭다운 선택
        await selectTop10FilterDropdown(page, '금액 기준', anomaly.type, menuName);

        // 필터 반영 확인: 해당 월 데이터가 테이블에 나타날 때까지 대기
        try {
            await page.waitForFunction(
                month => {
                    const rows = document.querySelectorAll('table tbody tr');
                    return rows.length > 0 &&
                        [...rows].some(r => r.textContent.includes(month));
                },
                anomaly.month,
                { timeout: 10000 }
            );
        } catch {
            await page.waitForTimeout(2000); // 폴백 대기
        }

        // 파일명: {filePrefix}월별트렌드_이상치_{YYYYMM}_{차대구분}.xlsx
        const monthSlug = anomaly.month.replace('-', ''); // "2025-04" → "202504"
        const saveName  = `월별트렌드_이상치_${monthSlug}_${anomaly.type}.xlsx`;
        const savePath  = path.join(resultsDir, `${filePrefix}${saveName}`);

        // TOP 10 섹션의 '엑셀 다운로드' 버튼 클릭
        try {
            const top10Section = page.locator('section, div').filter({
                hasText: /월별 거래처 Top 10|월별 거래처 TOP 10/,
            }).last();

            await page.waitForSelector('button:has-text("엑셀 다운로드")', {
                state: 'visible', timeout: 15000,
            });

            const downloadPromise = page.waitForEvent('download');

            const top10Btn = top10Section.locator('button:has-text("엑셀 다운로드")').first();
            if (await top10Btn.count() > 0) {
                await top10Btn.click();
            } else {
                // fallback: 화면 내 마지막 다운로드 버튼
                const allBtns = page.locator('button:has-text("엑셀 다운로드")');
                await allBtns.nth(await allBtns.count() - 1).click();
            }

            const download    = await downloadPromise;
            const downloadedPath = await download.path();

            // EBUSY 재시도 (OneDrive 동기화 대비)
            for (let attempt = 1; attempt <= 5; attempt++) {
                try {
                    fs.copyFileSync(downloadedPath, savePath);
                    console.log(`[${menuName}] 저장 완료: ${path.basename(savePath)}`);
                    break;
                } catch (e) {
                    if (e.code === 'EBUSY' && attempt < 5) {
                        await new Promise(r => setTimeout(r, attempt * 1000));
                    } else throw e;
                }
            }
        } catch (e) {
            console.log(`[경고] [${menuName}] ${anomaly.month} ${anomaly.type} 다운로드 실패: ${e.message}`);
        }

        await page.waitForTimeout(1000); // 다음 이상치 처리 전 안정화 대기
    }
}

// ─── /ai-analysis 업로드 영역별 파일 주입 헬퍼 ──────────────────────────────
// areaIndex: 0 = 분개장(필수), 1 = 계정별원장(선택)
// 전략 1: 업로드 영역 레이블 근처의 input 탐색 → 전략 2: nth(areaIndex) 폴백
async function uploadFileToZone(page, config, filePath, areaIndex, areaLabel, menuName) {
    console.log(`[${menuName}] ${areaLabel} 파일 업로드 시작: ${path.basename(filePath)}`);

    const fileInputSelector = config.selectors.fileUploadInput || 'input[type="file"]';

    // 전략 1: 업로드 버튼(uploadButton)이 설정된 경우 — 첫 번째 영역(분개장)만 해당
    if (areaIndex === 0 && config.selectors.uploadButton) {
        try {
            const [fileChooser] = await Promise.all([
                page.waitForEvent('filechooser', { timeout: 10000 }),
                page.click(config.selectors.uploadButton),
            ]);
            await fileChooser.setFiles(filePath);
            console.log(`[${menuName}] ${areaLabel} 파일 선택 완료 (fileChooser 방식).`);
            await page.waitForTimeout(1000);
            return;
        } catch {
            console.log(`[${menuName}] fileChooser 방식 실패, 직접 주입 방식으로 전환합니다.`);
        }
    }

    // 전략 2: nth(areaIndex) waitFor — DOM 렌더링 완료 후 hidden input 포함 직접 주입
    try {
        const nthInput = page.locator(fileInputSelector).nth(areaIndex);
        await nthInput.waitFor({ state: 'attached', timeout: 5000 });
        await nthInput.setInputFiles(filePath);
        console.log(`[${menuName}] ${areaLabel} 파일 주입 완료 (setInputFiles nth=${areaIndex}).`);
        await page.waitForTimeout(1000);
        return;
    } catch { /* 전략 3으로 */ }

    // 전략 3: 계정별원장 섹션의 드롭존 클릭 → filechooser 이벤트
    console.log(`[${menuName}] ${areaLabel} nth input 미발견 — 드롭존 클릭 방식으로 전환합니다.`);
    const dropZoneSelectors = [
        // 섹션 헤더 기준으로 내부 드롭존 탐색 (가장 정확)
        `div:has(> h2:has-text("계정별원장"), > h3:has-text("계정별원장"), > strong:has-text("계정별원장")) div[role="button"]`,
        // 텍스트 기준 드롭존 직접 탐색
        `div:has-text("당기 계정별원장 파일 업로드"):not(:has(*))`,   // 자식 없는 최하위 div
        `p:has-text("당기 계정별원장 파일 업로드")`,
        // nth 기반 폴백
        `[role="button"]:nth(${areaIndex})`,
        `[tabindex="0"]:nth(${areaIndex})`,
    ];
    let triggered = false;
    for (const sel of dropZoneSelectors) {
        try {
            const zone = page.locator(sel).first();
            if (await zone.count().catch(() => 0) === 0) continue;
            const [fileChooser] = await Promise.all([
                page.waitForEvent('filechooser', { timeout: 6000 }),
                zone.click(),
            ]);
            await fileChooser.setFiles(filePath);
            console.log(`[${menuName}] ${areaLabel} 파일 선택 완료 (드롭존 클릭: "${sel}").`);
            await page.waitForTimeout(1000);
            triggered = true;
            break;
        } catch { /* 다음 셀렉터 */ }
    }

    // 전략 4: 드롭존 클릭 후 단일 input에 직접 주입 (React가 input을 재활용하는 경우)
    if (!triggered) {
        console.log(`[${menuName}] ${areaLabel} filechooser 미발생 — 드롭존 클릭 후 nth(0) 주입 시도.`);
        const fallbackZone = page.locator(
            'section:has-text("계정별원장 파일 업로드"), ' +
            'div:has-text("계정별원장 파일 업로드 (선택사항)")'
        ).last();
        try {
            await fallbackZone.click({ timeout: 3000 });
            await page.waitForTimeout(500);
            await page.locator(fileInputSelector).nth(0).setInputFiles(filePath);
            console.log(`[${menuName}] ${areaLabel} 파일 주입 완료 (드롭존 클릭 후 nth=0 주입).`);
            triggered = true;
        } catch { /* 최종 실패 */ }
    }

    if (!triggered) {
        console.log(`[경고] [${menuName}] ${areaLabel} 업로드 실패 — 건너뜁니다.`);
    }
}

// ─── AI 분석 대시보드 복귀 헬퍼 ──────────────────────────────────────────────
// 분석 완료 후 [초기화면으로] 버튼을 클릭하여 대시보드로 복귀.
// 성공 시 true 반환 (분개장 세션 유지). 실패 시 false 반환 (세션 끊김 처리 필요).
// ★ 절대 browser.back() 또는 URL 재접속을 사용하지 않는다 — 세션(업로드 데이터)이 소실됨.
async function returnToAiDashboard(page, menuName) {
    const btnSel = 'button:has-text("초기화면으로"), a:has-text("초기화면으로")';
    try {
        await page.waitForSelector(btnSel, { state: 'visible', timeout: 5000 });
        await page.click(btnSel);
        await page.waitForTimeout(1500);
        // 대시보드 확인: 분석 카드 중 하나가 나타나면 복귀 성공
        await page.waitForSelector(
            'text=전표분석, text=일반사항 분석, text=월별 트렌드 분석, text=공휴일전표',
            { timeout: 10000 }
        );
        console.log(`[${menuName}] ✓ [초기화면으로] 복귀 완료 — 분개장 세션 유지 중.`);
        return true;
    } catch (e) {
        console.log(`[경고] [${menuName}] [초기화면으로] 실패 또는 대시보드 미복귀: ${e.message}`);
        console.log(`[${menuName}] 세션 끊김 감지 — 다음 메뉴에서 파일 재업로드 예정.`);
        return false;
    }
}

// ─── /ai-analysis 엔드포인트 메뉴 핸들러 ────────────────────────────────────
// task_list 컬럼:
//   업로드파일 / 분개장파일  → 더존 분개장 파일 경로 (필수, companyDir 기준 상대경로 또는 절대경로)
//   계정별원장파일 / 원장파일 → 당기 계정별원장 파일 경로 (선택)
//   결과파일명               → 다운로드 저장 파일명 (생략 시 menuName_결과 사용)
// skipUpload: true면 파일 업로드 단계를 건너뜀 (이전 메뉴에서 세션이 유지된 경우).
async function handleAiAnalysisMenu(page, menu, config, companyDir, rawDataDir, filePrefix, skipUpload = false) {
    const { menuName, tasks } = menu;
    console.log(`\n=== [메뉴 진입] ${menuName} (AI 분석)${skipUpload ? ' [업로드 생략 — 세션 유지]' : ''} ===`);

    // ── HOST2_월별트렌드 계열: 이상치 감지 핸들러로 분기 ─────────────────────────
    const isMonthlyTrend = /월별트렌드/.test(getBaseMenuName(menuName));
    // 대시보드에서 클릭할 분석 카드 UI 레이블
    const uiCardLabel = getMenuUiLabel(menuName, config);

    for (const task of tasks) {

        // ── 1~4. 파일 업로드 (세션이 없을 때만 수행) ─────────────────────────
        if (!skipUpload) {
            // 1. 분개장 파일 경로 확인 (필수)
            const journalRelPath = String(
                task['업로드파일'] ?? task['분개장파일'] ?? task['파일명'] ?? config.aiJournalFileName ?? config.uploadFileName ?? ''
            );
            if (!journalRelPath) {
                console.log(`[${menuName}] 분개장 파일이 지정되지 않았습니다. 건너뜁니다.`);
                continue;
            }
            const journalFilePath = path.isAbsolute(journalRelPath)
                ? journalRelPath
                : path.join(companyDir, journalRelPath);
            if (!fs.existsSync(journalFilePath)) {
                console.log(`[${menuName}] 분개장 파일이 존재하지 않습니다: ${journalFilePath}`);
                continue;
            }

            // 2. 분개장 업로드 (첫 번째 업로드 영역)
            await uploadFileToZone(page, config, journalFilePath, 0, '분개장', menuName);

            console.log(`[${menuName}] 분개장 처리 대기 중...`);
            try {
                await page.waitForSelector('text=데이터 건수', { timeout: 30000 });
                console.log(`[${menuName}] 분개장 업로드 완료.`);
            } catch {
                console.log(`[${menuName}] 분개장 완료 신호 미감지 — 5초 추가 대기합니다.`);
                await page.waitForTimeout(5000);
            }

            // 3. 계정별원장 파일 업로드 (두 번째 업로드 영역, 선택)
            const ledgerRelPath = String(task['계정별원장파일'] ?? task['원장파일'] ?? config.aiLedgerFileName ?? '');
            if (ledgerRelPath) {
                const ledgerFilePath = path.isAbsolute(ledgerRelPath)
                    ? ledgerRelPath
                    : path.join(companyDir, ledgerRelPath);

                if (fs.existsSync(ledgerFilePath)) {
                    await uploadFileToZone(page, config, ledgerFilePath, 1, '계정별원장', menuName);

                    console.log(`[${menuName}] 계정별원장 처리 대기 중...`);
                    try {
                        await page.waitForSelector('text=시트 수', { timeout: 30000 });
                        console.log(`[${menuName}] 계정별원장 업로드 완료.`);
                    } catch {
                        console.log(`[${menuName}] 계정별원장 완료 신호 미감지 — 5초 추가 대기합니다.`);
                        await page.waitForTimeout(5000);
                    }
                } else {
                    console.log(`[${menuName}] 계정별원장 파일 미발견 (건너뜀): ${ledgerFilePath}`);
                }
            } else {
                console.log(`[${menuName}] 계정별원장 파일 미지정. 분개장만 업로드합니다.`);
            }

            // 4. 업로드 완료 후 분석 카드 대시보드 대기
            try {
                await page.waitForSelector(
                    'text=전표분석, text=일반사항 분석, text=공휴일전표',
                    { timeout: 15000 }
                );
                console.log(`[${menuName}] 분석 카드 대시보드 전환 완료.`);
            } catch {
                console.log(`[${menuName}] 분석 카드 미감지 — 현재 화면에서 계속 진행합니다.`);
            }
        }

        // ── 5. 대시보드에서 분석 카드 클릭 ──────────────────────────────────
        // 월별트렌드는 handleMonthlyTrendAnalysis 내부에서 직접 처리하므로 제외.
        if (!isMonthlyTrend) {
            try {
                const card = page.locator(
                    `text="${uiCardLabel}", h2:has-text("${uiCardLabel}"), ` +
                    `h3:has-text("${uiCardLabel}"), div:has-text("${uiCardLabel}")`
                ).first();
                if (await card.count() > 0) {
                    await card.waitFor({ state: 'visible', timeout: 8000 });
                    await card.evaluate(n => {
                        n.removeAttribute('target');
                        n.closest('a')?.removeAttribute('target');
                    });
                    await card.click();
                    await page.waitForTimeout(2000);
                    console.log(`[${menuName}] ✓ 분석 카드 "${uiCardLabel}" 진입 완료.`);
                }
            } catch (e) {
                console.log(`[경고] [${menuName}] 분석 카드 클릭 실패, 현재 화면에서 계속합니다: ${e.message}`);
            }
        }

        // ── 6. 메뉴 유형별 분석 실행 ─────────────────────────────────────────
        if (isMonthlyTrend) {
            await handleMonthlyTrendAnalysis(page, menu, config, companyDir, rawDataDir, filePrefix);
            return; // task 반복 불필요 — 핸들러 내부에서 전체 처리
        }

        // ── 6b. Google AI Studio: "AI 심층 분석 시작" → 태스크별 카드 처리 ──
        try {
            const aiStartBtn = page.locator('button:has-text("AI 심층 분석 시작")').first();
            if (await aiStartBtn.count().catch(() => 0) > 0) {
                await aiStartBtn.click();
                await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
                await page.waitForTimeout(1500);
                console.log(`[${menuName}] AI 심층 분석 대시보드 진입 완료.`);
                if (tasks.some(t => t['작업명'])) {
                    await handleGoogleAiAnalysis(page, menu, config, rawDataDir, filePrefix);
                    return;
                }
            }
        } catch { /* 없으면 기존 플로우 */ }

        // 일반 AI 분석: 결과 다운로드
        const downloadBtnSelector = config.selectors.excelDownloadBtn || 'button:has-text("결과 다운로드")';
        const outputFileName = String(task['결과파일명'] ?? task['파일명'] ?? `${menuName}_결과`);
        await handleDownloadAndSave(page, downloadBtnSelector, outputFileName, rawDataDir, menuName, filePrefix);
    }
}

// ─── Google AI Studio 심층 분석: 태스크별 카드 클릭/다운로드 ──────────────────
async function handleGoogleAiAnalysis(page, menu, config, resultsDir, filePrefix) {
    const { menuName, tasks } = menu;

    const TASK_UI_MAP = {
        '일반사항분석': '일반사항 분석',
        '공휴일전표': '공휴일전표',
        '상대계정분석': '상대계정 분석',
        '적요적합성분석': '적요 적합성 분석',
        '시각화분석': '시각화 분석',
        '월별트렌드분석': '월별 트렌드 분석',
        '현금흐름분석': '현금 흐름 분석',
    };

    const returnToDashboard = async () => {
        const btnSel = 'button:has-text("초기화면으로"), a:has-text("초기화면으로")';
        try {
            await page.waitForSelector(btnSel, { state: 'visible', timeout: 5000 });
            await page.click(btnSel);
            await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
            await page.waitForTimeout(1500);
            return true;
        } catch { return false; }
    };

    for (const task of tasks) {
        const taskName  = String(task['작업명']   ?? '').trim();
        const account   = String(task['계정과목'] ?? '').trim();
        const direction = String(task['거래방향'] ?? '').trim();
        if (!taskName) continue;

        const uiLabel = TASK_UI_MAP[taskName] ?? taskName;
        const logTag  = `${taskName}${account ? `/${account}` : ''}`;
        console.log(`\n--- [${menuName} / ${logTag}] 처리 시작 ---`);

        // 카드 클릭
        try {
            const card = page.locator(`text="${uiLabel}"`).first();
            if (await card.count().catch(() => 0) === 0) {
                console.log(`  [경고] "${uiLabel}" 카드 미발견.`);
                continue;
            }
            await card.waitFor({ state: 'visible', timeout: 8000 });
            await card.click();
            await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
            await page.waitForTimeout(1500);
            console.log(`  ✓ 카드 "${uiLabel}" 진입 완료`);
        } catch (e) {
            console.log(`  [경고] 카드 클릭 실패: ${e.message}`);
            continue;
        }

        // 계정과목 필터 (combobox 없으면 일반 input 검색창으로 폴백)
        if (account) {
            try {
                const comboSel = config.selectors.accountCombobox || 'button[role="combobox"]';
                let combo = page.locator(comboSel).first();
                if (await combo.count().catch(() => 0) === 0) {
                    combo = page.locator('input[type="search"], input[placeholder]').first();
                }
                if (await combo.count().catch(() => 0) > 0) {
                    await combo.click();
                    await page.waitForTimeout(300);
                    await page.keyboard.press('Control+A');
                    await page.keyboard.press('Backspace');
                    await page.keyboard.type(account, { delay: 50 });
                    await page.waitForTimeout(500);
                    await page.keyboard.press('Enter');
                    console.log(`  계정과목: ${account}`);

                    // 계정 변경 후 데이터 리로드 대기
                    // 요약 통계(총 차변 합계 등) 또는 networkidle로 갱신 확인
                    try {
                        await page.waitForLoadState('networkidle', { timeout: 8000 });
                    } catch { /* networkidle 미감지 시 폴백 */ }
                    // 차트/테이블이 해당 계정 데이터로 갱신될 때까지 추가 대기
                    try {
                        await page.waitForFunction(
                            acc => {
                                const inputs = document.querySelectorAll('input[placeholder], input[type="search"]');
                                for (const el of inputs) {
                                    if (el.value && el.value.includes(acc)) return true;
                                }
                                // 드롭다운/선택된 값 텍스트로도 확인
                                const body = document.body.innerText;
                                return body.includes('총 차변') || body.includes('총 분석 월 수');
                            },
                            account,
                            { timeout: 5000 }
                        );
                    } catch { /* 폴백: 고정 대기 */ }
                    await page.waitForTimeout(1500);
                }
            } catch { /* 필터 없으면 무시 */ }
        }

        // 거래방향 라디오
        if (direction) await clickRadioByLabel(page, direction, '거래방향');

        // 분석 실행 버튼 자동 클릭 제거 — 카드 진입 후 결과 자동 표시됨 (AI API 불필요 호출 방지)

        // 결과 대기 (최대 5분) + 다운로드
        // 월별트렌드분析: 버튼 1(금액추이)·2(건수)만 다운로드. 버튼 3(Top10)은 이상치 루프에서 처리.
        const isMonthlyTrendTask = taskName === '월별트렌드분析';
        const dlSel = 'button:has-text("결과 다운로드"), button:has-text("엑셀 다운로드")';
        try {
            await page.waitForSelector(dlSel, { state: 'visible', timeout: 300000 });

            let dlBtns = await page.locator('button:has-text("엑셀 다운로드")').all();
            if (dlBtns.length === 0) dlBtns = await page.locator('button:has-text("결과 다운로드")').all();

            const safeTask = taskName.replace(/[\\/?*[\]:]/g, '_');
            const safeAcc  = account   ? `_${account}`   : '';
            const safeDir  = direction ? `_${direction}`  : '';
            const baseName = `${filePrefix}${safeTask}${safeAcc}${safeDir}`;

            // 월별트렌드분析은 버튼 1·2만 다운로드 (Top10 버튼 제외)
            const downloadCount = isMonthlyTrendTask ? Math.min(dlBtns.length, 2) : dlBtns.length;

            for (let i = 0; i < downloadCount; i++) {
                const suffix   = dlBtns.length > 1 ? `_${i + 1}` : '';
                const savePath = path.join(resultsDir, `${baseName}${suffix}.xlsx`);

                try {
                    await dlBtns[i].scrollIntoViewIfNeeded();

                    const dl = await new Promise(resolve => {
                        const timer = setTimeout(() => {
                            page.off('download', onDl);
                            resolve(null);
                        }, 15000);
                        function onDl(download) {
                            clearTimeout(timer);
                            page.off('download', onDl);
                            resolve(download);
                        }
                        page.on('download', onDl);
                        dlBtns[i].click().catch(() => {
                            clearTimeout(timer);
                            page.off('download', onDl);
                            resolve(null);
                        });
                    });

                    if (!dl) {
                        console.log(`  [건너뜀] 버튼 ${i + 1}: 다운로드 이벤트 없음.`);
                        continue;
                    }

                    const dlPath = await dl.path();
                    for (let attempt = 1; attempt <= 5; attempt++) {
                        try {
                            fs.copyFileSync(dlPath, savePath);
                            console.log(`  ✓ 저장 완료: ${path.basename(savePath)}`);
                            break;
                        } catch (e) {
                            if (e.code === 'EBUSY' && attempt < 5) await new Promise(r => setTimeout(r, attempt * 1000));
                            else throw e;
                        }
                    }
                    await page.waitForTimeout(500);
                } catch (e) {
                    console.log(`  [경고] 버튼 ${i + 1} 다운로드 실패: ${e.message}`);
                }
            }
        } catch (e) {
            console.log(`  [경고] 결과 다운로드 실패: ${e.message}`);
        }

        // ── 월별트렌드분析: 이상치 감지 → Top10(3번 버튼) 조건부 다운로드 ──────────
        // 이상치(급증/급감 모두)가 있는 달에 대해서만 월+차대변 필터 설정 후 3번 버튼 클릭
        if (isMonthlyTrendTask) {
            try {
                console.log(`\n  [이상치분析] Pre-Scan 시작 — 계정: ${account}`);
                const monthlyData = await extractMonthlyAmountsFromPage(page, taskName);

                if (monthlyData.length === 0) {
                    console.log(`  [이상치분析] 월별 데이터 추출 실패 — Top10 생략`);
                } else {
                    const anomalies = detectMonthlyAnomalies(monthlyData, 0.3);

                    if (anomalies.length === 0) {
                        console.log(`  [이상치분析] 이상치 없음 (평균 ±30%) — Top10 다운로드 생략`);
                    } else {
                        console.log(`  [이상치분析] ${anomalies.length}건 감지 → Top10(3번 버튼) 다운로드 시작`);
                        anomalies.forEach((a, i) =>
                            console.log(`    ${i + 1}. ${a.month} [${a.type}] ${a.amount.toLocaleString()} (평균 ${Math.round(a.avg).toLocaleString()})`)
                        );

                        // 3번 버튼(index 2) = 월별 거래처 Top10 엑셀 다운로드
                        const top10Btn = page.locator('button:has-text("엑셀 다운로드")').nth(2);
                        if (await top10Btn.count() === 0) {
                            console.log(`  [경고] Top10 버튼(3번)을 찾지 못했습니다 — 이상치 다운로드 생략`);
                        } else {
                            for (const anomaly of anomalies) {
                                const saveName = `${filePrefix}월별트렌드_${account}_${anomaly.month}_${anomaly.type}.xlsx`;
                                const savePath = path.join(resultsDir, saveName);
                                console.log(`\n  → [${anomaly.month}][${anomaly.type}] Top10 처리 중…`);

                                try {
                                    // 월 + 금액 기준(차변/대변) 드롭다운 설정
                                    await selectTop10FilterDropdown(page, '월', anomaly.month, taskName);
                                    await selectTop10FilterDropdown(page, '금액 기준', anomaly.type, taskName);

                                    // 테이블 갱신 대기
                                    try {
                                        await page.waitForSelector(
                                            '[class*="loading"], [class*="spinner"], [aria-busy="true"]',
                                            { state: 'hidden', timeout: 5000 }
                                        ).catch(() => {});
                                        await page.waitForFunction(
                                            () => document.querySelectorAll('table tbody tr').length >= 1,
                                            { timeout: 8000 }
                                        );
                                        await page.waitForTimeout(800);
                                    } catch {
                                        await page.waitForTimeout(2000);
                                    }

                                    // 3번 버튼 클릭 + 다운로드
                                    await top10Btn.scrollIntoViewIfNeeded();
                                    const dl = await new Promise(resolve => {
                                        const timer = setTimeout(() => {
                                            page.off('download', onDl);
                                            resolve(null);
                                        }, 15000);
                                        function onDl(download) {
                                            clearTimeout(timer);
                                            page.off('download', onDl);
                                            resolve(download);
                                        }
                                        page.on('download', onDl);
                                        top10Btn.click().catch(() => {
                                            clearTimeout(timer);
                                            page.off('download', onDl);
                                            resolve(null);
                                        });
                                    });

                                    if (!dl) {
                                        console.log(`  [건너뜀] ${anomaly.month} ${anomaly.type} — 다운로드 이벤트 없음`);
                                        continue;
                                    }

                                    const dlPath = await dl.path();
                                    for (let attempt = 1; attempt <= 5; attempt++) {
                                        try {
                                            fs.copyFileSync(dlPath, savePath);
                                            console.log(`  ✓ 저장: ${saveName}`);
                                            break;
                                        } catch (e) {
                                            if (e.code === 'EBUSY' && attempt < 5) {
                                                await new Promise(r => setTimeout(r, attempt * 1000));
                                            } else throw e;
                                        }
                                    }
                                    await page.waitForTimeout(500);
                                } catch (e) {
                                    console.log(`  [경고] ${anomaly.month} ${anomaly.type} Top10 실패: ${e.message}`);
                                }
                            }
                        }
                    }
                }
            } catch (e) {
                console.log(`  [경고] 월별 이상치 처리 실패: ${e.message}`);
            }
        }

        // 대시보드 복귀
        const returned = await returnToDashboard();
        if (!returned) console.log(`  [경고] 대시보드 복귀 실패.`);
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
        let aiSessionActive   = false;  // /ai-analysis 분개장 세션 유지 여부

        for (const menu of config.menus) {
            const menuName = menu.menuName;
            const endpoint = getMenuEndpoint(menuName, config);
            const targetUrl = `${baseUrl}${endpoint}`;

            // 엔드포인트가 바뀔 때만 페이지 이동
            if (currentEndpoint !== endpoint) {
                console.log(`\n[라우팅] ${menuName} → ${targetUrl}`);
                await page.goto(targetUrl, { waitUntil: 'networkidle', timeout: 60000 });
                currentEndpoint = endpoint;
                analysisUploadDone = false;
                aiSessionActive   = false; // 페이지 이동 시 분개장 세션 초기화
                await page.waitForTimeout(1000);
            }

            if (endpoint === '/ai-analysis') {
                // ── AI 분석: 세션 유지 시 업로드 생략, 완료 후 [초기화면으로] 복귀 ──
                await handleAiAnalysisMenu(
                    page, menu, config, companyDir, resultsDir, filePrefix,
                    /* skipUpload = */ aiSessionActive
                );

                // 분석 완료 후 [초기화면으로] 버튼으로 대시보드 복귀 (세션 유지)
                const returned = await returnToAiDashboard(page, menuName);
                if (returned) {
                    aiSessionActive = true;  // 다음 메뉴는 업로드 생략 가능
                } else {
                    aiSessionActive = false; // 세션 끊김 → 다음 메뉴에서 재업로드
                }

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
                // 카드 클릭 헬퍼: button/a → 텍스트 포함 요소 순으로 탐색
                const findMenuHandle = async () => {
                    for (const sel of [
                        `button:has-text("${uiLabel}")`,
                        `a:has-text("${uiLabel}")`,
                        `[role="button"]:has-text("${uiLabel}")`,
                    ]) {
                        const loc = page.locator(sel).first();
                        if (await loc.count().catch(() => 0) > 0) return loc;
                    }
                    // 폴백: 텍스트를 정확히 포함하는 요소
                    const h = await page.$(`text="${uiLabel}"`).catch(() => null)
                        ?? await page.$(`h2:has-text("${uiLabel}"), h3:has-text("${uiLabel}"), span:has-text("${uiLabel}"), div:has-text("${uiLabel}")`).catch(() => null);
                    return h;
                };

                // 카드 클릭 실행 (최대 2회 시도)
                const clickMenuCard = async (handle) => {
                    if (!handle) return;
                    // target="_blank" 제거 후 클릭
                    try {
                        await page.evaluate(node => {
                            node.removeAttribute?.('target');
                            node.closest?.('a')?.removeAttribute('target');
                        }, handle.elementHandle ? await handle.elementHandle() : handle);
                    } catch { /* 무시 */ }
                    await (handle.click ? handle.click() : page.click(handle));
                    // 페이지 전환 안정화 대기
                    await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
                    await page.waitForTimeout(1500);
                };

                let menuHandle = await findMenuHandle();

                // 카드를 못 찾으면 '뒤로가기' 또는 URL 재이동 후 재탐색
                if (!menuHandle) {
                    try {
                        const backBtn = await page.waitForSelector(
                            'button:has-text("뒤로가기"), a:has-text("뒤로가기")',
                            { state: 'visible', timeout: 5000 }
                        );
                        console.log(`[안내] "${uiLabel}" 카드 미발견 → '뒤로가기' 클릭으로 메인 화면 복귀합니다.`);
                        await backBtn.click();
                        await page.waitForLoadState('networkidle', { timeout: 15000 });
                        await page.waitForTimeout(500);
                    } catch {
                        console.log(`[안내] '뒤로가기' 버튼 미발견 → ${targetUrl}로 URL 재이동합니다.`);
                        await page.goto(targetUrl, { waitUntil: 'networkidle', timeout: 60000 });
                        await page.waitForTimeout(1000);
                    }
                    menuHandle = await findMenuHandle();
                }

                if (menuHandle) {
                    await clickMenuCard(menuHandle);
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
