'use strict';
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { spawnSync } = require('child_process');
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
    '이자비용적정성test':     '상세 거래 검색',   // 작업명 컬럼 사용 시트 → 동일 UI 카드 진입
    '특수관계자거래':         '상세 거래 검색',   // 거래처명 단독 검색 시트 → 동일 UI 카드 진입
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
// 우선순위: raw_data/previous/ → raw_data/ flat (하위 호환)
function findPrevYearLedgerFile(rawDataDir) {
    if (!fs.existsSync(rawDataDir)) return null;

    // 1순위: previous/ 서브폴더
    const prevDir = path.join(rawDataDir, 'previous');
    if (fs.existsSync(prevDir)) {
        const found = fs.readdirSync(prevDir).find(f => {
            const ext = path.extname(f).toLowerCase();
            return (ext === '.xlsx' || ext === '.xls') && !f.startsWith('~$');
        });
        if (found) return path.join(prevDir, found);
    }

    // 2순위: flat raw_data/ (하위 호환 — 파일명에 '전기' 포함)
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
    // "아니오" / "아니요" 두 표기 모두 대응
    const DISMISS_SEL = 'button:has-text("아니오"), button:has-text("아니요")';

    if (prevYearFile) {
        // 전기 파일이 있을 때: "네, 전기 데이터도 업로드하겠습니다" 클릭 후 파일 주입
        try {
            const yesPrevBtn = await page.waitForSelector(
                'button:has-text("네")',
                { state: 'visible', timeout: 10000 }   // 5000 → 10000
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
            // 다이얼로그가 화면에 남아 있으면 강제로 닫아 이후 메뉴 클릭 차단 방지
            try {
                const dismissBtn = page.locator(DISMISS_SEL).first();
                if (await dismissBtn.isVisible({ timeout: 3000 })) {
                    await dismissBtn.click();
                    await page.waitForTimeout(800);
                    console.log('[업로드] 전기 데이터 다이얼로그 강제 닫기 완료');
                }
            } catch { /* 다이얼로그가 없으면 무시 */ }
        }
    } else {
        // 전기 파일이 없을 때: "아니오/아니요, 당기만 분석하겠습니다" 클릭
        try {
            const skipPrevBtn = await page.waitForSelector(
                DISMISS_SEL,
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

// ─── 다운로드 → workbook 단일 시트에 누적 추가 헬퍼 ─────────────────────────
// 그룹 내 여러 계정 결과를 sheetName 시트 하나에 누적(append). 첫 호출 시
// 헤더 행을 포함해 시트를 생성하고, 이후 호출부터는 헤더를 제외한 데이터 행만
// 기존 시트 마지막 행 다음부터 추가한다. 파일은 저장하지 않음.
async function downloadAndAppendToSheet(page, downloadBtnSelector, sheetName, workbook, menuName) {
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

    let destSheet = workbook.getWorksheet(safeSheetName);
    const isNewSheet = !destSheet;
    if (isNewSheet) destSheet = workbook.addWorksheet(safeSheetName);

    let destRowNumber = isNewSheet ? 0 : destSheet.rowCount;
    srcSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        // 두 번째 이후 호출에서는 소스의 헤더 행(1행)을 건너뛰고 데이터 행만 추가
        if (!isNewSheet && rowNumber === 1) return;
        destRowNumber += 1;
        const destRow = destSheet.getRow(destRowNumber);
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            destRow.getCell(colNumber).value = cell.value;
        });
        destRow.commit();
    });
    console.log(`[${menuName}] '${safeSheetName}' 시트에 결과 추가 완료 (현재 ${destSheet.rowCount}행).`);
}

// ─── 리스 완전성 자동 연동: lease_filter.py 실행 ─────────────────────────────
function runLeaseFilter(companyName, noFilter = false, outputDir = null) {
    const leaseScript = path.join(__dirname, '..', 'lease_analyzer', 'lease_filter.py');
    if (!fs.existsSync(leaseScript)) {
        console.log(`[리스완전성] lease_filter.py 를 찾을 수 없습니다: ${leaseScript}`);
        return;
    }
    const args = [leaseScript, '--company', companyName];
    if (noFilter) args.push('--no-filter');
    if (outputDir) {
        const outPath = path.join(outputDir, `리스완전성_${companyName}.xlsx`);
        args.push('--output', outPath);
    }
    console.log(`\n[리스완전성] lease_filter.py 자동 실행 (회사: ${companyName}${noFilter ? ', 키워드필터 생략' : ''})`);
    const result = spawnSync('python', args, {
        encoding: 'utf8',
        env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
        timeout: 120000,
    });
    if (result.stdout) process.stdout.write(result.stdout);
    if (result.stderr) {
        // lxml UserWarning 등 무해한 경고는 생략, 실제 오류만 출력
        const errLines = result.stderr.split('\n').filter(l =>
            l.trim() && !l.includes('UserWarning') && !l.includes('pkg_resources')
        );
        if (errLines.length) console.error('[lease_filter]', errLines.join('\n'));
    }
    if (result.status === 0) {
        console.log('[리스완전성] lease_filter.py 완료');
    } else if (result.status !== null) {
        console.log(`[리스완전성] lease_filter.py 종료 코드: ${result.status}`);
    } else if (result.error) {
        console.log(`[리스완전성] lease_filter.py 실행 실패: ${result.error.message}`);
    }
}

// ─── 은행조회서완전성 자동 연동: bank_confirmation_filter.py 실행 ───────────
function runBankConfirmation(filePath) {
    const script = path.join(__dirname, '..', 'bank_confirmation_filter.py');
    if (!fs.existsSync(script)) {
        console.log(`[은행조회서완전성] bank_confirmation_filter.py 를 찾을 수 없습니다: ${script}`);
        return;
    }
    console.log(`\n[은행조회서완전성] bank_confirmation_filter.py 자동 실행`);
    const result = spawnSync('python', [script, '--file', filePath], {
        encoding: 'utf8',
        env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
        timeout: 60000,
    });
    if (result.stdout) process.stdout.write(result.stdout);
    if (result.stderr) {
        const errLines = result.stderr.split('\n').filter(l =>
            l.trim() && !l.includes('UserWarning') && !l.includes('pkg_resources')
        );
        if (errLines.length) console.error('[은행조회서완전성]', errLines.join('\n'));
    }
    if (result.status === 0) {
        console.log('[은행조회서완전성] bank_confirmation_filter.py 완료');
    } else if (result.status !== null) {
        console.log(`[은행조회서완전성] bank_confirmation_filter.py 종료 코드: ${result.status}`);
    } else if (result.error) {
        console.log(`[은행조회서완전성] bank_confirmation_filter.py 실행 실패: ${result.error.message}`);
    }
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

        // '분석옵션' 컬럼에 '단일시트'/'통합' 입력 시, 그룹 내 모든 계정 결과를 계정별 시트가 아닌
        // 하나의 시트로 합쳐서 저장 (다운로드 결과에 이미 '계정과목' 컬럼이 포함되어 있음)
        const groupOptionRaw  = String(groupTasks[0]['분석옵션'] ?? groupTasks[0]['분석 옵션'] ?? '').trim();
        const mergeIntoOneSheet = /단일시트|통합|merge/i.test(groupOptionRaw);
        const combinedSheetName = taskName.substring(0, 31).replace(/[\\/?*[\]:]/g, '_');

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

            // 계정과목·거래처명 모두 없으면 처리 불가 → 건너뜀
            if (!accountName && !vendorName) {
                console.log(`[${taskName}] 계정과목·거래처명 모두 없음 — 건너뜁니다.`);
                continue;
            }
            // 계정과목 없이 거래처명만 있을 때: 거래처명 단독 검색 모드
            const vendorOnlyMode = !accountName && !!vendorName;
            // 시트명: 계정과목 우선, 없으면 거래처명
            const sheetLabel = accountName || vendorName;
            if (vendorOnlyMode) {
                console.log(`\n--- [${taskName} / 거래처: ${vendorName}] 거래처명 단독 검색 ---`);
            } else {
                console.log(`\n--- [${taskName} / ${accountName}] 처리 시작 ---`);
            }

            try {
                // 1. 콤보박스가 보일 때까지 대기
                await page.waitForSelector(comboSel, { state: 'visible', timeout: 10000 });
                await page.waitForTimeout(500);

                // 2. 계정과목 입력 (계정과목이 있을 때만)
                if (accountName) {
                    await page.click(comboSel);           // 1차 클릭: 콤보박스 활성화
                    await page.waitForTimeout(400);
                    const acctInput = page.locator('input[placeholder*="계정"], input[placeholder*="입력"]').first();
                    if (await acctInput.isVisible({ timeout: 800 }).catch(() => false)) {
                        await acctInput.click();
                    } else {
                        await page.click(comboSel);
                    }
                    await page.waitForTimeout(300);
                    await page.keyboard.type(accountName, { delay: 50 });
                    await page.waitForTimeout(700);

                    // 드롭다운 옵션 선택: 정확한 텍스트 일치 우선 → 첫 번째 옵션 클릭 → ArrowDown+Enter 폴백
                    let acctSelected = false;
                    try {
                        const exactOpt = page.locator(`[role="option"]:text-is("${accountName}")`).first();
                        if (await exactOpt.isVisible({ timeout: 1500 })) {
                            await exactOpt.click();
                            acctSelected = true;
                            console.log(`  계정과목: ${accountName} (정확히 일치하는 옵션 클릭)`);
                        }
                    } catch {}
                    if (!acctSelected) {
                        try {
                            const firstOpt = page.locator('[role="option"]').first();
                            if (await firstOpt.isVisible({ timeout: 1000 })) {
                                const optText = await firstOpt.textContent().catch(() => '');
                                await firstOpt.click();
                                acctSelected = true;
                                console.log(`  계정과목: ${accountName} → 첫 번째 옵션 클릭 ("${optText?.trim()}")`);
                            }
                        } catch {}
                    }
                    if (!acctSelected) {
                        await page.keyboard.press('ArrowDown');
                        await page.waitForTimeout(200);
                        await page.keyboard.press('Enter');
                        console.log(`  계정과목: ${accountName} (ArrowDown+Enter 폴백)`);
                    }
                    await page.waitForTimeout(400);
                }

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
                        // Escape 대신 Tab: Escape는 드롭다운을 닫으며 이전 선택값을 복원하므로 사용 금지
                        await page.keyboard.press('Tab');
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
                        // Escape 대신 Tab: 이전 선택값 복원 방지
                        await page.keyboard.press('Tab');
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

                // 7. 검색 — 드롭다운이 열려 있으면 클릭이 씹히므로 Escape로 먼저 닫고 버튼 클릭
                await page.keyboard.press('Escape');
                await page.waitForTimeout(300);
                // 검색 직전 스크린샷 — 계정과목이 실제로 선택됐는지 확인용 (첫 계정만)
                if (ti === 0 && gi === 0) {
                    await page.screenshot({ path: `graphy/debug_before_search_${safeFileName}.png` }).catch(() => {});
                    console.log(`  [디버그] 검색 전 스크린샷 저장: debug_before_search_${safeFileName}.png`);
                }
                const searchBtn = page.getByRole('button', { name: '검색', exact: true });
                await searchBtn.waitFor({ state: 'visible', timeout: 10000 });
                await searchBtn.click();
                console.log(`  검색 버튼 클릭 완료`);
                await page.waitForTimeout(2000);

                // 8. 다운로드 → 그룹 workbook에 시트 추가
                const downloadSel = config.selectors.excelDownloadBtn || 'button:has-text("결과 다운로드")';
                const dlVisible = await page.locator(downloadSel).waitFor({ state: 'visible', timeout: 10000 }).then(() => true).catch(() => false);
                if (!dlVisible) {
                    // 결과 없음: 현재 버튼 목록 + 스크린샷으로 원인 파악
                    const btns = await page.locator('button').allTextContents().catch(() => []);
                    const safeName = sheetLabel.replace(/[^\w가-힣]/g, '_');
                    console.log(`  [안내] '${sheetLabel}' 검색 결과 없음. 버튼 목록: [${btns.map(t => t.trim()).filter(Boolean).join(' | ')}]`);
                    await page.screenshot({ path: `graphy/debug_no_result_${safeName}.png` }).catch(() => {});
                } else if (mergeIntoOneSheet) {
                    await downloadAndAppendToSheet(page, downloadSel, combinedSheetName, groupBook, menuName);
                } else {
                    await downloadAndAddSheet(page, downloadSel, sheetLabel, groupBook, menuName);
                }

            } catch (e) {
                console.log(`[경고] [${taskName} / ${sheetLabel}] 처리 중 오류 (다음 계정으로 진행): ${e.message}`);
            }

            // 9. 다음 태스크가 있으면: [초기화] 버튼 클릭 → 폼 안정화 → 다음 계정 입력 준비
            const isLastOverall = gi === allGroups.length - 1 && ti === groupTasks.length - 1;
            if (!isLastOverall) {
                const resetSel = config.selectors.resetButton || 'button:has-text("초기화")';
                try {
                    await page.waitForSelector(resetSel, { state: 'visible', timeout: 5000 });
                    await page.locator(resetSel).last().click(); // last(): 폼 하단 메인 초기화 버튼 우선
                    await page.waitForTimeout(1000); // 폼 초기화 안정화 대기
                    await page.waitForSelector(comboSel, { state: 'visible', timeout: 10000 });
                    console.log(`  → [초기화] 완료, 다음 계정 입력 준비`);
                } catch (e) {
                    console.log(`[경고] 초기화 버튼 클릭 실패 (다음 계정으로 진행): ${e.message}`);
                }
            }
        }

        // 10. 그룹 파일 저장 (OneDrive EBUSY 재시도 포함)
        // xlsx 규격상 시트가 0개이면 Excel이 손상으로 인식하므로 빈 안내 시트 추가
        if (groupBook.worksheets.length === 0) {
            const emptySheet = groupBook.addWorksheet('결과없음');
            emptySheet.getCell('A1').value = '검색 결과가 없습니다.';
            console.log(`[${taskName}] 검색 결과 없음 — '결과없음' 시트를 추가하여 파일을 저장합니다.`);
        }
        // 임시 파일에 먼저 쓴 뒤 rename — OneDrive EBUSY 회피
        const tempGroupPath = groupFilePath + '.tmp';
        await groupBook.xlsx.writeFile(tempGroupPath);
        let groupSaved = false;
        for (let attempt = 1; attempt <= 15; attempt++) {
            try {
                if (fs.existsSync(groupFilePath)) fs.unlinkSync(groupFilePath);
                fs.renameSync(tempGroupPath, groupFilePath);
                console.log(`[${taskName}] 그룹 파일 저장 완료: ${path.basename(groupFilePath)}`);
                groupSaved = true;
                break;
            } catch (e) {
                if ((e.code === 'EBUSY' || e.code === 'EPERM') && attempt < 15) {
                    const wait = Math.min(attempt, 5);
                    console.log(`[${taskName}] 파일 잠금 감지, ${wait}초 후 재시도... (${attempt}/15)`);
                    await new Promise(r => setTimeout(r, wait * 1000));
                } else {
                    console.log(`[${taskName}][경고] 파일 저장 최종 실패 — 임시 파일 유지: ${path.basename(tempGroupPath)}`);
                    break;
                }
            }
        }

        // 11. 리스 완전성 시나리오이면 lease_filter.py 자동 연동
        //     task_list 시트 '분석옵션' 컬럼에 'no-filter' 또는 '전건' 입력 시 키워드 필터 생략
        if (/리스/.test(taskName) && config.companyName) {
            const optionRaw = String(groupTasks[0]['분석옵션'] ?? groupTasks[0]['분석 옵션'] ?? '').trim();
            // Playwright로 내려받은 계정은 이미 리스 후보 선별 완료 → 키워드 필터 불필요
            // task_list에 '필터' 또는 'filter'라고 명시한 경우에만 키워드 필터 적용
            const useFilter = /^필터$|^filter$/i.test(optionRaw);
            runLeaseFilter(config.companyName, !useFilter, resultsDir);
        }

        // 12. 은행조회서완전성 시나리오이면 bank_confirmation_filter.py 자동 연동
        //     저장된 그룹 파일에 금융기관명 컬럼 + 조회서목록 요약 시트를 추가
        if (/은행/.test(taskName)) {
            runBankConfirmation(groupFilePath);
        }
    }
}

// ─── 다운로드 & 저장 헬퍼 ────────────────────────────────────────────────────
async function handleDownloadAndSave(page, downloadBtnSelector, targetName, rawDataDir, menuName, filePrefix = '', btnTimeout = 30000) {
    console.log(`[${menuName}] 결과 다운로드 버튼 대기 중...`);
    await page.waitForSelector(downloadBtnSelector, { state: 'visible', timeout: btnTimeout });

    console.log(`[${menuName}] 다운로드를 진행합니다.`);
    const downloadPromise = page.waitForEvent('download', { timeout: btnTimeout });
    await page.click(downloadBtnSelector);
    let download, downloadPath;
    try {
        download = await downloadPromise;
        downloadPath = await download.path();
    } catch (dlErr) {
        console.log(`[${menuName}][경고] 다운로드 이벤트 타임아웃 — 데이터 없음으로 건너뜁니다. (${targetName})`);
        return;
    }
    console.log(`[${menuName}] 임시 다운로드 캡처 완료.`);

    // 마스터 파일 병합 대상
    const _baseMN = menuName.replace(/^host_?\d+_?/i, '');
    // 벤포드는 차트 보존을 위해 마스터 병합 제외 → 계정별 개별 파일로 저장
    const MASTER_MERGE_MENUS = ['상세 거래 검색', '총계정원장조회', '총계정원장'];
    if (MASTER_MERGE_MENUS.includes(_baseMN)) {
        const baseFileName = (_baseMN === '상세 거래 검색') ? '상세거래검색.xlsx'
            : (_baseMN === '벤포드법칙분석') ? '벤포드법칙분석.xlsx'
            : '총계정원장.xlsx';
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
        const safeTarget = targetName.replace(/[\\/?*[\]:]/g, '_');
        const finalName = safeTarget.startsWith(filePrefix) ? safeTarget : `${filePrefix}${safeTarget}`;
        const finalPath = path.join(rawDataDir, finalName.endsWith('.xlsx') ? finalName : `${finalName}.xlsx`);
        // OneDrive EBUSY 회피: 기존 파일 삭제 후 copyFile
        let copied = false;
        for (let attempt = 1; attempt <= 15; attempt++) {
            try {
                // 기존 파일이 잠겨 있으면 삭제 먼저 시도
                if (fs.existsSync(finalPath)) {
                    try { fs.unlinkSync(finalPath); } catch { /* 잠금 시 copyFileSync가 덮어쓰기 시도 */ }
                }
                fs.copyFileSync(downloadPath, finalPath);
                copied = true;
                break;
            } catch (e) {
                if ((e.code === 'EBUSY' || e.code === 'EPERM') && attempt < 15) {
                    const wait = Math.min(attempt, 5);
                    console.log(`[${menuName}] 파일 잠금 감지, ${wait}초 후 재시도... (${attempt}/15)`);
                    await new Promise(r => setTimeout(r, wait * 1000));
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

    // ── 상세검색_시나리오 시트: 전용 핸들러로 위임
    // '작업명' 컬럼이 있는 시트는 이름에 관계없이 동일 핸들러로 처리
    const hasTaskName = tasks.some(t => '작업명' in t && t['작업명'] !== null);
    if (base === '상세검색_시나리오' || hasTaskName) {
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
            const accountCombo = (await combos.count().catch(() => 0)) > 0 ? combos.first() : null;
            if (accountCombo) {
                await accountCombo.click();
                await page.waitForTimeout(500);
                await page.keyboard.press('Control+A');
                await page.keyboard.press('Backspace');
                await page.keyboard.type(accountName, { delay: 50 });
                await page.waitForTimeout(500);
                await page.keyboard.press('Enter');
                await page.waitForTimeout(800); // 계정 선택 후 금액기준열 콤보 렌더링 대기
                console.log(`  ✓ 계정과목 '${accountName}' 선택`);
            }

            // 2) 금액 기준열 선택 (task에 '금액기준열' / '금액 기준열' / '기준열' 컬럼이 있으면 적용)
            const amountCol = String(
                task['금액기준열'] ?? task['금액 기준열'] ?? task['기준열'] ?? ''
            ).trim();
            if (amountCol) {
                let set = false;
                // 계정 선택 후 콤보 상태 재평가
                const comboCount = await combos.count().catch(() => 0);

                // 전략 1: native <select> — 차변/대변/코드 옵션을 포함한 select를 탐색
                try {
                    const selects = page.locator('select');
                    const selCount = await selects.count().catch(() => 0);
                    for (let si = 0; si < selCount && !set; si++) {
                        const opts = await selects.nth(si).locator('option').allTextContents().catch(() => []);
                        if (opts.some(o => ['차변', '대변', '코드'].includes(o.trim()))) {
                            await selects.nth(si).selectOption({ label: amountCol });
                            set = true;
                        }
                    }
                    if (set) console.log(`  ✓ 금액 기준열 '${amountCol}' 선택`);
                } catch { /* fallthrough */ }

                // 전략 2: button[role="combobox"] 두 번째 항목 (커스텀 드롭다운)
                if (!set && comboCount >= 2) {
                    await combos.nth(1).click();
                    await page.waitForTimeout(400);
                    try {
                        await page.locator(`[role="option"]:has-text("${amountCol}")`).first().click({ timeout: 3000 });
                        set = true;
                        console.log(`  ✓ 금액 기준열 '${amountCol}' 선택`);
                    } catch {
                        await page.keyboard.press('Escape');
                    }
                    await page.waitForTimeout(400);
                }

                if (!set) console.log(`  [경고] 금액 기준열 '${amountCol}' 설정 실패 — 기본값(코드) 유지`);
            }

            // 3) 분석 시작 클릭 → AI 감사인 의견 생성 완료까지 대기
            // AI 의견은 서버에서 별도 생성되며 DOM text에 나타나지 않음.
            // 차트/테이블은 ~15s에 안정화되지만 AI 생성은 최대 90s 소요.
            // 전략: 분석 클릭 후 최소 90초 대기(AI 생성 시간 확보) + 텍스트 안정화 확인
            const baselineLen = await page.evaluate(() => (document.body.innerText || '').length).catch(() => 0);
            await page.click('button:has-text("분석 시작")');
            const clickedAt = Date.now();
            console.log(`  ✓ 분석 시작 클릭 — AI 감사인 의견 생성 대기 중... (기준 텍스트: ${baselineLen}자)`);

            // 1단계: 차트/테이블 로드 완료까지 대기 (텍스트 안정화)
            let prevLen = -1;
            let stableCount = 0;
            let analysisStarted = false;
            let chartLoaded = false;
            for (let attempt = 0; attempt < 24 && !chartLoaded; attempt++) {
                await page.waitForTimeout(5000);
                try {
                    const currentLen = await page.evaluate(() => (document.body.innerText || '').length);
                    console.log(`  [대기] ${(attempt + 1) * 5}초 경과 — 페이지 텍스트: ${currentLen}자`);
                    if (prevLen === -1) { prevLen = currentLen; continue; }
                    if (currentLen !== prevLen) { analysisStarted = true; stableCount = 0; }
                    else {
                        stableCount++;
                        if ((analysisStarted && stableCount >= 2) || stableCount >= 6) chartLoaded = true;
                    }
                    prevLen = currentLen;
                } catch { stableCount = 0; }
            }

            // 2단계: AI 의견 생성 대기 — 분석 시작 후 30초 추가 대기 (AI가 20초 내 생성됨)
            const elapsed2 = Date.now() - clickedAt;
            const aiWait = Math.max(0, 30000 - elapsed2);
            if (aiWait > 0) {
                console.log(`  [대기] AI 의견 생성 중... (${Math.ceil(aiWait / 1000)}초 대기)`);
                await page.waitForTimeout(aiWait);
            }
            console.log(`  ✓ AI 생성 대기 완료 (총 ${Math.round((Date.now() - clickedAt) / 1000)}초 경과)`);

            await page.waitForTimeout(1000);  // 렌더링 안정화

            // 4) 결과 다운로드 (벤포드 결과 섹션의 "엑셀 다운로드" 버튼)
            // 동일 계정에 차변/대변을 모두 분석하는 경우가 있어 파일명에 분석 기준(차변/대변)을 포함
            const targetName = task['파일명']
                ? String(task['파일명'])
                : `${accountName}${amountCol ? `_${amountCol}` : ''}`;
            const dlBtn = config.selectors.benfordDownloadBtn || 'button:has-text("엑셀 다운로드")';
            await handleDownloadAndSave(page, dlBtn, targetName, rawDataDir, menuName, filePrefix);

            // 5) 다음 계정 처리 전 여유 대기 (API 연속 호출 부하 방지)
            await page.waitForTimeout(2000);
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

            // 드롭다운 옵션 확인 — 없으면 해당 계정이 원장에 없는 것
            const dropdownOptions = page.locator('[role="option"]');
            const optionCount = await dropdownOptions.count().catch(() => 0);
            if (optionCount === 0) {
                console.log(`[${menuName}][건너뜀] '${accountName}' — 계정 없음 (드롭다운 옵션 미발견)`);
                await page.keyboard.press('Escape');
                continue;
            }

            await page.keyboard.press('Enter');
            await page.waitForTimeout(500);

            if (IS_LEDGER_MENU) {
                // 검색 버튼 없음 — 네트워크 요청 완료 후 다운로드
                // 고정 대기 대신 networkidle로 실제 데이터 로딩 완료를 확인
                console.log(`[${accountName}] 데이터 로딩 대기 중 (networkidle)...`);
                try {
                    await page.waitForLoadState('networkidle', { timeout: 30000 });
                } catch {
                    console.log(`[${accountName}] 네트워크 대기 타임아웃 — 추가 3초 대기 후 진행.`);
                    await page.waitForTimeout(3000);
                }
                await page.waitForTimeout(500); // 렌더링 안정화 대기
                const ledgerTarget = String(task['파일명'] ?? `${accountName}_${base}`);
                await handleDownloadAndSave(page, 'button:has-text("엑셀 다운로드")', ledgerTarget, rawDataDir, menuName, filePrefix);

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
        console.log(`[${menuName}] 분석 중... (계정 수가 많아 시간이 소요될 수 있습니다)`);
        // 분석 완료까지 networkidle 대기 (최대 3분)
        await page.waitForLoadState('networkidle', { timeout: 180000 }).catch(() => {});
        await page.waitForTimeout(1000);
        const fileName = String(task['파일명'] ?? base);
        // 다운로드 버튼 대기도 120초로 설정
        await handleDownloadAndSave(page, 'button:has-text("엑셀 다운로드")', fileName, rawDataDir, menuName, filePrefix, 120000);

    } else if (['외상매출매입상계', '외상매출/매입 상계 거래처 분석'].includes(base)) {
        console.log(`\n--- [${menuName}] 처리 시작 ---`);
        const task = tasks[0] ?? {};
        await page.click('button:has-text("상계 거래처 분석 시작")');
        await page.waitForTimeout(1000);
        const fileName = String(task['파일명'] ?? base);
        await handleDownloadAndSave(page, 'button:has-text("엑셀 다운로드")', fileName, rawDataDir, menuName, filePrefix);

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

            // 2. 기준월 버튼 선택 (3월/6월/9월/12월)
            const baseMonth = task['기준월'];
            if (baseMonth) {
                await clickRadioByLabel(page, `${baseMonth}월`, '기준월');
                await page.waitForTimeout(1000);  // 기간 변경 후 재집계 대기
            }

            // 3. 금액 유형 라디오 (차변만 / 대변만 / 차변+대변 모두)
            if (amountType) {
                await clickRadioByLabel(page, amountType, '금액 유형');
                await page.waitForTimeout(1000);
            }

            // 4. 비교표 다운로드
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
        // 이미 대시보드에 있으면 성공 처리 (이중 호출 방지)
        const hasHome = await page.locator('button,a').filter({ hasText: '초기화면으로' }).first().isVisible({ timeout: 1000 }).catch(() => false);
        if (!hasHome) {
            const onBoard = await page.locator('button,a').filter({ hasText: '상세보기' }).first().isVisible({ timeout: 1500 }).catch(() => false);
            if (onBoard) { console.log(`[${menuName}] ✓ 이미 대시보드 — 세션 유지 중.`); return true; }
        }
        await page.waitForSelector(btnSel, { state: 'visible', timeout: 8000 });
        await page.click(btnSel);
        await page.waitForTimeout(1500);
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
    // 한자析(U+6790) 변형 키 자동 추가: 엑셀에서 한자로 입력된 작업명 대응
    Object.keys(TASK_UI_MAP).forEach(k => { TASK_UI_MAP[k.replace(/석/g, String.fromCodePoint(0x6790))] = TASK_UI_MAP[k]; });

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

        // 카드 클릭: 첫 단어(순수 한글)로 컨테이너 탐색 → 상세보기 버튼 또는 헤딩 클릭
        try {
            const keyword = uiLabel.split(' ')[0];
            let clicked = false;
            try {
                const btn = page.locator('div,section,li,article,[class*="card"]')
                    .filter({ has: page.locator('h1,h2,h3,h4,p,span').filter({ hasText: keyword }) })
                    .locator('button,a').filter({ hasText: '상세보기' }).first();
                if (await btn.count().catch(() => 0) > 0) { await btn.click(); clicked = true; }
            } catch {}
            if (!clicked) {
                const heading = page.locator('h1,h2,h3,h4').filter({ hasText: keyword }).first();
                if (await heading.count().catch(() => 0) > 0) { await heading.click(); clicked = true; }
            }
            if (!clicked) {
                clicked = await page.evaluate((kw) => {
                    for (const el of document.querySelectorAll('h1,h2,h3,h4,p,button,span')) {
                        if ((el.textContent||'').trim().includes(kw) && el.offsetParent) { el.click(); return true; }
                    }
                    return false;
                }, keyword).catch(() => false);
            }
            if (!clicked) { console.log(`  [경고] "${uiLabel}" 카드 미발견.`); continue; }
            await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
            await page.waitForTimeout(1500);
            console.log(`  ✓ 카드"${uiLabel}" 진입 완료`);
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
                    await page.waitForTimeout(500);
                    // Popover+Command UI 감지: 트리거 클릭 후 [cmdk-input] 존재 여부 확인
                    const cmdInput = page.locator('[cmdk-input]').first();
                    const hasCmdInput = await cmdInput.count().catch(() => 0) > 0;
                    if (hasCmdInput) {
                        // 신규 UI: Popover 내 CommandInput에 직접 타이핑
                        await cmdInput.fill('');
                        await cmdInput.type(account, { delay: 50 });
                    } else {
                        // 구 UI: 키보드로 직접 타이핑
                        await page.keyboard.press('Control+A');
                        await page.keyboard.press('Backspace');
                        await page.keyboard.type(account, { delay: 50 });
                    }
                    await page.waitForTimeout(800);
                    // 드롭다운 항목 클릭: cmdk-item 우선 → [코드]계정명 패턴 탐색 → locator 폴백
                    let acctSelected = false;
                    // 전략 1: [cmdk-item] — [코드] 제거 후 텍스트만 비교, 정확 일치 우선
                    if (!acctSelected) {
                        try {
                            const acctName = account.replace(/^\[\d+\]\s*/, '').trim();
                            const allItems = await page.locator('[cmdk-item]').all();
                            let bestItem = null;
                            for (const item of allItems) {
                                const txt = (await item.textContent().catch(() => '')).trim();
                                const nameOnly = txt.replace(/^\[\d+\]\s*/, '');
                                if (nameOnly === acctName || nameOnly.startsWith(acctName + '(') || nameOnly.startsWith(acctName + ' ')) { bestItem = { el: item, txt }; break; }
                            }
                            if (!bestItem) {
                                for (const item of allItems) {
                                    const txt = (await item.textContent().catch(() => '')).trim();
                                    const nameOnly = txt.replace(/^\[\d+\]\s*/, '');
                                    if (nameOnly.includes(acctName) && !nameOnly.includes('(' + acctName + ')')) { bestItem = { el: item, txt }; break; }
                                }
                            }
                            if (bestItem) {
                                await bestItem.el.click();
                                acctSelected = true;
                                console.log('  계정과목: ' + account + ' → cmdk-item 클릭: ' + bestItem.txt.substring(0, 40));
                            }
                        } catch {}
                    }
                    // 전략 2: page.evaluate로 [코드]계정명 패턴 탐색 (구 UI 호환)
                    if (!acctSelected) {
                        const clickedOpt = await page.evaluate((acc) => {
                            const codePattern = /^\[\d+\]/;
                            const tags = ["li","div","button","span","p","a"];
                            for (const tag of tags) {
                                for (const el of document.querySelectorAll(tag)) {
                                    const txt = (el.textContent || "").trim();
                                    if (txt.includes(acc) && (codePattern.test(txt) || txt === acc) && el.offsetParent !== null && el.children.length <= 2) {
                                        el.click();
                                        return txt;
                                    }
                                }
                            }
                            return null;
                        }, account);
                        if (clickedOpt) {
                            acctSelected = true;
                            console.log("  계정과목: " + account + " → 드롭다운 클릭: " + clickedOpt.substring(0,40));
                        }
                    }
                    // 전략 3: locator로 [코드] 패턴 항목 클릭
                    if (!acctSelected) {
                        try {
                            const opt = page.locator("div, li, button, span").filter({ hasText: new RegExp("\\[\\d+\\].*" + account.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")) }).first();
                            if (await opt.count() > 0 && await opt.isVisible({ timeout: 1500 })) {
                                const txt = await opt.textContent().catch(() => "");
                                await opt.click();
                                acctSelected = true;
                                console.log("  계정과목: " + account + " → locator 클릭: " + (txt||"").trim().substring(0,40));
                            }
                        } catch {}
                    }
                    if (!acctSelected) {
                        await page.keyboard.press("ArrowDown");
                        await page.waitForTimeout(300);
                        await page.keyboard.press("Enter");
                        console.log("  계정과목: " + account + " (ArrowDown+Enter 폴백)");
                    }
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

        // 상대계정분석은 계정/방향 선택 후 분석 실행 버튼을 눈러야 결과가 나타남
        if (taskName && (taskName.includes('상대계정') || menuName.includes('상대계정'))) {
            try {
                const allBtns = await page.locator('button').allTextContents().catch(() => []);
                console.log('  [디버그] 상대계정분석 버튼 목록: [' + allBtns.map(t => t.trim()).filter(Boolean).join(' | ') + ']');
                const runBtn = page.locator('button').filter({ hasText: /분석.{0,3}실행|실행/ }).first();
                if (await runBtn.count().catch(() => 0) > 0) {
                    const btnText = await runBtn.textContent().catch(() => '');
                    await runBtn.click();
                    console.log('  ✓ 분석 실행 클릭: ' + btnText.trim());
                    // 분析 완료 감지: 버튼 변화 폴링 (10초 x 6회)
                    for (let chk = 1; chk <= 6; chk++) {
                        await page.waitForTimeout(10000);
                        const btnsAfter = await page.locator('button').allTextContents().catch(() => []);
                        console.log('  [디버그] 분석 후 ' + (chk*10) + 's 버튼: [' + btnsAfter.map(t => t.trim()).filter(Boolean).join(' | ') + ']');
                        const hasDownload = btnsAfter.some(t => /다운로드|download/i.test(t));
                        if (hasDownload) { console.log('  [디버그] 다운로드 버튼 감지 — 대기 종료'); break; }
                    }
                    const debugSsPath = require('path').join(__dirname, '..', 'graphy', 'debug_상대계정_' + (account || 'unknown') + '.png');
                    await page.screenshot({ path: debugSsPath, fullPage: false }).catch(() => {});
                    console.log('  [디버그] 스크린샷: ' + debugSsPath);
                } else {
                    console.log('  [경고] 분석 실행 버튼을 찾지 못했습니다.');
                }
            } catch (e) { console.log('  [경고] 분석 실행 버튼 처리 오류: ' + e.message); }
        }

        // 결과 대기 (최대 5분) + 다운로드
        // 월별트렌드분석: 버튼 1(금액추이)·2(건수)만 다운로드. 버튼 3(Top10)은 이상치 루프에서 처리.
        const isMonthlyTrendTask = taskName === '월별트렌드분석';
        const dlSel = 'button:has-text("결과 다운로드"), button:has-text("엑셀 다운로드")';
        try {
            const isRelatedAccountTask = taskName && taskName.includes('상대계정');
            if (isRelatedAccountTask) {
                // 분析 완료 후 DOM에 삽입되는 다운로드 버튼으로 완료 감지 (opacity-0이라 attached 사용)
                await page.locator('button:has-text("요약 엑셀 다운로드"), button:has-text("엑셀 다운로드"), button:has-text("결과 다운로드")').first().waitFor({ state: 'attached', timeout: 300000 });
                await page.waitForTimeout(1000);
                console.log('  [상대계정] 결과 테이블 감지 — force click으로 다운로드 시도');
            } else {
                await page.waitForSelector(dlSel, { state: 'visible', timeout: 300000 });
            }

            let dlBtns = await page.locator('button:has-text("엑셀 다운로드")').all();
            if (dlBtns.length === 0) dlBtns = await page.locator('button:has-text("결과 다운로드")').all();
            if (dlBtns.length === 0) dlBtns = await page.locator('text=요약 엑셀 다운로드').all();
            if (dlBtns.length === 0) dlBtns = await page.locator('button:has-text("요약 엑셀 다운로드")').all();

            const clickOptions = (taskName && taskName.includes('상대계정')) ? { force: true } : {};
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
                        dlBtns[i].click(clickOptions).catch(() => {
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
                            const seen = new Set();
                            const uniqueAnomalies = anomalies.filter(a => {
                                const key = `${a.month}_${a.type}`;
                                if (seen.has(key)) return false;
                                seen.add(key);
                                return true;
                            });
                            for (const anomaly of uniqueAnomalies) {
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
        const loginFormVisible = await page.locator(emailSelector).isVisible().catch(() => false);

        if (loginFormVisible && config.credentials?.userId) {
            try {
                // 3초 내에 실제로 입력 가능 상태인지 재확인 (false positive 방지)
                await page.waitForSelector(emailSelector, { state: 'visible', timeout: 3000 });
                console.log(`[${companyName}] 로그인 폼 감지. 로그인을 진행합니다...`);
                const pwSelector = config.selectors.loginPassword || 'input[type="password"]';
                const loginBtnSelector = config.selectors.loginButton || 'button:has-text("로그인")';
                await page.fill(emailSelector, config.credentials.userId);
                await page.fill(pwSelector, config.credentials.userPassword ?? '');
                await page.click(loginBtnSelector);
                console.log(`[${companyName}] 로그인 완료. 화면 전환 대기 중...`);
                await page.waitForTimeout(2000);
            } catch {
                console.log(`[${companyName}] 기존 세션 감지 (로그인 폼 미활성). 로그인을 생략합니다.`);
            }
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

                // 카드 셀렉터 전략: 정확한 텍스트 일치 우선 → 역할 기반 폴백
                // div:has-text() 는 상위 컨테이너 전체를 매칭해 오클릭을 유발하므로 사용하지 않음.
                const findMenuHandle = async () => {
                    const strategies = [
                        // 1순위: :text-is() — 요소 텍스트가 정확히 uiLabel인 것만
                        () => page.locator(`:text-is("${uiLabel}")`).first(),
                        // 2순위: getByText exact
                        () => page.getByText(uiLabel, { exact: true }).first(),
                        // 3순위: heading role 정확 일치
                        () => page.getByRole('heading', { name: uiLabel, exact: true }).first(),
                        // 4순위: h 태그 정규식 정확 일치
                        () => page.locator('h2, h3, h4').filter({ hasText: new RegExp(`^${uiLabel.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`) }).first(),
                        // 5순위: 역할 기반 (has-text 부분 일치 — 낮은 우선순위)
                        () => page.locator(`button:has-text("${uiLabel}")`).first(),
                        () => page.locator(`a:has-text("${uiLabel}")`).first(),
                        () => page.locator(`[role="button"]:has-text("${uiLabel}")`).first(),
                    ];
                    for (const getFn of strategies) {
                        try {
                            const loc = getFn();
                            if (await loc.count().catch(() => 0) > 0) return loc;
                        } catch { /* 다음 전략 */ }
                    }
                    return null;
                };

                // 카드 클릭 실행
                const clickMenuCard = async (loc) => {
                    if (!loc) return;
                    await loc.evaluate(n => {
                        n.removeAttribute?.('target');
                        n.closest?.('a')?.removeAttribute('target');
                    }).catch(() => {});
                    await loc.click();
                    await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
                    await page.waitForTimeout(1500);
                };

                let menuHandle = await findMenuHandle();

                // 카드를 못 찾으면 '뒤로가기' 또는 URL 재이동 후 재탐색
                if (!menuHandle) {
                    let wentBack = false;
                    try {
                        const backBtn = await page.waitForSelector(
                            'button:has-text("뒤로가기"), a:has-text("뒤로가기")',
                            { state: 'visible', timeout: 5000 }
                        );
                        console.log(`[안내] "${uiLabel}" 카드 미발견 → '뒤로가기' 클릭으로 메인 화면 복귀합니다.`);
                        await backBtn.click();
                        // networkidle 타임아웃이 catch로 전파되지 않도록 분리
                        await page.waitForLoadState('networkidle', { timeout: 10000 }).catch(() => {});
                        await page.waitForTimeout(1000);
                        wentBack = true;
                    } catch {
                        // '뒤로가기' 버튼 자체를 못 찾은 경우에만 URL 재이동
                    }
                    if (!wentBack) {
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
