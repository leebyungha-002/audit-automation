'use strict';
// ─── 이자비용 적정성 데이터 추출 (전역 공통 모듈) ───────────────────────────────
// 실행: node interest_expense_extractor.js <회사명>
//   예: node interest_expense_extractor.js dae_il
//       node interest_expense_extractor.js braintree
//
// 경로 규칙 (루트 = audit-automation/)
//   시나리오 : ./<회사명>/task_list_<회사명>.xlsx
//   결과물   : ./<회사명>/results/<작업명>.xlsx
//   원장파일 : ./<회사명>/raw_data/*.xlsx|xls  (당기, ~$제외)
//   브라우저 : ./.browser_profile
//   환경변수 : ./.env

const path    = require('path');
const fs      = require('fs');
const ExcelJS = require('exceljs');

// ── CLI 인수 검증 ────────────────────────────────────────────────────────────
const COMPANY = (process.argv[2] ?? '').trim();
if (!COMPANY) {
    console.error('[오류] 회사명 인수가 필요합니다.');
    console.error('  사용법: node interest_expense_extractor.js <회사명>');
    console.error('  예시  : node interest_expense_extractor.js dae_il');
    process.exit(1);
}

const ROOT_DIR    = __dirname;                                          // audit-automation/
const COMPANY_DIR = path.join(ROOT_DIR, COMPANY);                      // ./<회사명>/
const TASK_LIST_PATH = path.join(COMPANY_DIR, `task_list_${COMPANY}.xlsx`);
const RESULTS_DIR    = path.join(COMPANY_DIR, 'results');
const PROFILE_DIR    = path.join(ROOT_DIR, '.browser_profile');

require('dotenv').config({ path: path.join(ROOT_DIR, '.env') });

const { initBrowser } = require(path.join(ROOT_DIR, 'shared_modules', 'index'));

const BASE_URL = process.env.AUDIT_URL || 'http://127.0.0.1:8081';

const SHEET_PATTERNS = [
    'host1_이자비용적정성test',
    '이자비용적정성test',
    'host1_이자비용test',
    '이자비용test',
];

const COMBO_SEL    = 'button[role="combobox"]';
const SEARCH_SEL   = 'button:has-text("검색")';
const RESET_SEL    = 'button:has-text("초기화")';
const DOWNLOAD_SEL = 'button:has-text("결과 다운로드")';
const MENU_NAME    = '이자비용적정성';

// ─── 시트 로드 ────────────────────────────────────────────────────────────────
async function loadTasks() {
    if (!fs.existsSync(TASK_LIST_PATH))
        throw new Error(`시나리오 파일 없음: ${TASK_LIST_PATH}`);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TASK_LIST_PATH);

    let sheet = null;
    for (const name of SHEET_PATTERNS) {
        sheet = wb.getWorksheet(name);
        if (sheet) break;
    }
    if (!sheet) throw new Error(`이자비용적정성test 시트를 찾을 수 없습니다. (후보: ${SHEET_PATTERNS.join(', ')})`);

    let headers = null;
    const tasks = [];
    sheet.eachRow((row, rowNum) => {
        const vals = row.values; // vals[0] 은 항상 undefined
        if (rowNum === 1) {
            headers = vals.slice(1).map(h => (h != null ? String(h).trim() : null));
            return;
        }
        if (!headers) return;
        const task = {};
        headers.forEach((h, i) => { if (h) task[h] = vals[i + 1] ?? null; });
        if (String(task['계정과목'] ?? '').trim()) tasks.push(task);
    });

    console.log(`[시트] '${sheet.name}' → ${tasks.length}개 행 로드 완료`);
    return tasks;
}

// ─── 날짜 포맷 헬퍼 ──────────────────────────────────────────────────────────
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

// ─── 라디오 버튼 클릭 ────────────────────────────────────────────────────────
async function clickRadioByLabel(page, labelText, groupHint) {
    if (!labelText) return;
    const text    = String(labelText).trim();
    const exactRe = new RegExp(`^${text}$`);

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

// ─── 다운로드 → workbook 시트 추가 ──────────────────────────────────────────
async function downloadAndAddSheet(page, downloadBtnSelector, sheetName, workbook) {
    console.log(`[${MENU_NAME}] '${sheetName}' 결과 다운로드 대기 중...`);
    await page.waitForSelector(downloadBtnSelector, { state: 'visible', timeout: 30000 });

    const downloadPromise = page.waitForEvent('download');
    await page.click(downloadBtnSelector);
    const download     = await downloadPromise;
    const downloadPath = await download.path();
    console.log(`[${MENU_NAME}] 다운로드 캡처 완료.`);

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
    console.log(`[${MENU_NAME}] '${safeSheetName}' 시트 추가 완료.`);
}

// ─── 계정별원장 자동 탐색 ────────────────────────────────────────────────────
function findLedgerFile() {
    const rawDataDir = path.join(COMPANY_DIR, 'raw_data');
    if (!fs.existsSync(rawDataDir)) return null;
    const found = fs.readdirSync(rawDataDir).find(f => {
        const ext = path.extname(f).toLowerCase();
        return (ext === '.xlsx' || ext === '.xls') && !f.startsWith('~$') && !f.includes('전기');
    });
    return found ? path.join(rawDataDir, found) : null;
}

// ─── 메인 ─────────────────────────────────────────────────────────────────────
async function main() {
    console.log(`=== 이자비용 적정성 데이터 추출 시작 [${COMPANY}] ===`);
    console.log(`  시나리오: ${TASK_LIST_PATH}`);
    console.log(`  결과경로: ${RESULTS_DIR}\n`);

    const tasks = await loadTasks();
    if (!tasks.length) { console.log('[종료] 처리할 작업이 없습니다.'); return; }

    if (!fs.existsSync(RESULTS_DIR)) fs.mkdirSync(RESULTS_DIR, { recursive: true });

    const { browser, page } = await initBrowser(false, PROFILE_DIR);

    try {
        // ── 접속 ──────────────────────────────────────────────────────────────
        await page.goto(`${BASE_URL}/analysis`, { waitUntil: 'networkidle', timeout: 30000 });
        await page.waitForTimeout(1500);

        // ── 로그인 ────────────────────────────────────────────────────────────
        if (await page.locator('input[type="password"]').count().catch(() => 0) > 0) {
            const userId = process.env.USER_EMAIL    ?? '';
            const userPw = process.env.USER_PASSWORD ?? '';
            if (!userId || !userPw) throw new Error('.env 에 USER_EMAIL / USER_PASSWORD 가 없습니다.');
            await page.fill('input[type="email"]',    userId);
            await page.fill('input[type="password"]', userPw);
            await page.click('button:has-text("로그인"), #login-btn');
            await page.waitForLoadState('networkidle', { timeout: 15000 });
            await page.waitForTimeout(1500);
            console.log('[로그인] 완료');
        }

        // ── 계정별원장 업로드 (업로드 영역이 있는 경우만) ─────────────────────
        if (await page.locator('input[type="file"]').count().catch(() => 0) > 0) {
            const ledgerFile = findLedgerFile();
            if (ledgerFile) {
                console.log(`[업로드] ${path.basename(ledgerFile)}`);
                await page.setInputFiles('input[type="file"]', ledgerFile);
                try {
                    await page.waitForSelector('input[type="file"]', { state: 'hidden', timeout: 30000 });
                } catch { await page.waitForTimeout(5000); }

                try {
                    await page.waitForSelector('button:has-text("아니요")', { state: 'visible', timeout: 5000 });
                    await page.click('button:has-text("아니요")');
                    await page.waitForTimeout(1000);
                } catch { /* 전기 데이터 다이얼로그 없음 */ }

                try {
                    await page.waitForSelector('text=데이터 건수', { timeout: 60000 });
                    console.log('[업로드] 처리 완료');
                } catch { await page.waitForTimeout(5000); }
            } else {
                console.log('[업로드] 계정별원장 파일 없음 — 건너뜁니다.');
            }
        }

        // ── '상세 거래 검색' 카드 진입 ────────────────────────────────────────
        // 총계정원장 조회 카드 설명에 "상세 거래내역을 조회합니다"가 포함되어
        // has-text() 부분 일치 시 오매칭 발생 → 설명 고유 문구 우선 사용.
        if (await page.locator(COMBO_SEL).count().catch(() => 0) === 0) {
            const CARD = '상세 거래 검색';
            const strategies = [
                // 1순위: 카드 설명 고유 문구로 컨테이너 특정 (다른 카드와 중복 없음)
                () => page.locator('div, section, article, li')
                    .filter({ hasText: '거래처, 금액, 날짜, 적요' })
                    .first(),
                // 2순위: :text-is() — 요소 텍스트 정확 일치
                () => page.locator(':text-is("상세 거래 검색")').first(),
                // 3순위: heading role 정확 일치
                () => page.getByRole('heading', { name: CARD, exact: true }).first(),
                // 4순위: getByText exact
                () => page.getByText(CARD, { exact: true }).first(),
                // 5순위: h 태그 정규식 정확 일치
                () => page.locator('h2, h3, h4, h5').filter({ hasText: new RegExp(`^${CARD}$`) }).first(),
                // 역할 기반
                () => page.locator(`a:has-text("${CARD}")`).first(),
                () => page.locator(`button:has-text("${CARD}")`).first(),
                () => page.locator(`[role="button"]:has-text("${CARD}")`).first(),
            ];
            let entered = false;
            for (const getFn of strategies) {
                try {
                    const loc = getFn();
                    if (await loc.count().catch(() => 0) === 0) continue;
                    await loc.waitFor({ state: 'visible', timeout: 3000 });
                    await loc.evaluate(n => {
                        n.removeAttribute?.('target');
                        n.closest?.('a')?.removeAttribute('target');
                    }).catch(() => {});
                    await loc.click();
                    await page.waitForTimeout(2000);
                    // 올바른 폼이 열렸는지 확인
                    const onCorrectPage = await page.locator(COMBO_SEL).count().catch(() => 0) > 0
                        || await page.getByText('거래처, 금액, 날짜, 적요').count().catch(() => 0) > 0;
                    if (!onCorrectPage) {
                        console.log('  [재확인] 잘못된 화면 — 다음 전략으로');
                        continue;
                    }
                    console.log(`[카드] '${CARD}' 진입 완료`);
                    entered = true;
                    break;
                } catch { /* 다음 전략 */ }
            }
            if (!entered) console.log('[경고] 카드를 찾지 못했습니다. 현재 화면에서 계속합니다.');
        } else {
            console.log('[카드] 이미 상세 거래 검색 화면 — 카드 클릭 생략');
        }

        await page.waitForSelector(COMBO_SEL, { state: 'visible', timeout: 20000 });
        console.log('[검색폼] 준비 완료\n');

        // ── 작업명 기준 그룹화 ───────────────────────────────────────────────
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

        // ── 그룹별 처리 ─────────────────────────────────────────────────────
        for (let gi = 0; gi < allGroups.length; gi++) {
            const [taskName, groupTasks] = allGroups[gi];
            console.log(`\n=== [작업그룹: ${taskName}] ${groupTasks.length}개 계정 처리 시작 ===`);

            const groupBook     = new ExcelJS.Workbook();
            const safeFileName  = taskName.substring(0, 50).replace(/[\\/?*[\]:]/g, '_');
            const groupFilePath = path.join(RESULTS_DIR, `${safeFileName}.xlsx`);

            for (let ti = 0; ti < groupTasks.length; ti++) {
                const task        = groupTasks[ti];
                const accountName = String(task['계정과목'] ?? '').trim();
                const vendorName  = String(task['거래처명'] ?? '').trim();
                const description = String(task['적요']     ?? '').trim();
                const amountType  = String(task['금액 유형'] ?? task['금액유형'] ?? '').trim();
                const displayType = String(task['표시방식']  ?? task['표시 방식']  ?? '').trim();
                const startDateRaw = task['시작일'] ?? task['시작일자'] ?? task['기간시작'] ?? null;
                const endDateRaw   = task['종료일'] ?? task['종료일자'] ?? task['기간종료'] ?? null;

                if (!accountName) { console.log(`[${taskName}] 계정과목 없음 — 건너뜁니다.`); continue; }
                console.log(`\n--- [${taskName} / ${accountName}] 처리 시작 ---`);

                try {
                    // 1. 계정과목 combobox
                    // ※ Enter 키는 폼 자동 제출을 유발하므로 사용하지 않음.
                    //   드롭다운 항목이 있으면 클릭, 없으면 Escape 로 닫고 입력값 그대로 사용.
                    await page.waitForSelector(COMBO_SEL, { state: 'visible', timeout: 10000 });
                    await page.waitForTimeout(500);
                    await page.click(COMBO_SEL);
                    await page.waitForTimeout(300);
                    await page.keyboard.press('Control+A');
                    await page.keyboard.press('Backspace');
                    // 괄호 안 숫자 코드 추출 → 드롭다운 매칭 향상
                    // ex) "하나 장기운전자금4(29320)" → "29320"
                    const codeMatch  = accountName.match(/\((\d+)\)/) || accountName.match(/(\d+)\s*$/);
                    const searchText = codeMatch ? codeMatch[1] : accountName;
                    await page.keyboard.type(searchText, { delay: 50 });
                    await page.waitForTimeout(700);

                    const dropdownOption = page.locator('[role="option"], [role="listbox"] li, ul[role="listbox"] li').first();
                    if (await dropdownOption.count().catch(() => 0) > 0) {
                        await dropdownOption.click({ timeout: 2000 }).catch(() => {});
                        await page.waitForTimeout(300);
                        console.log(`  계정과목 선택 (검색어: "${searchText}"): ${accountName}`);
                    } else {
                        await page.keyboard.press('Escape');
                        await page.waitForTimeout(300);
                        console.log(`[경고] 계정과목 드롭다운 미발견 (검색어: "${searchText}") — 빈 상태로 진행`);
                    }

                    // 2. 거래처명 combobox
                    try {
                        const vendorCombo = page.locator(COMBO_SEL).nth(1);
                        await vendorCombo.click();
                        await page.waitForTimeout(200);
                        await page.keyboard.press('Control+A');
                        await page.keyboard.press('Backspace');
                        if (vendorName) {
                            await page.keyboard.type(vendorName, { delay: 50 });
                            await page.waitForTimeout(700);
                            const vendorOption = page.locator('[role="option"], [role="listbox"] li, ul[role="listbox"] li').first();
                            if (await vendorOption.count().catch(() => 0) > 0) {
                                await vendorOption.click({ timeout: 2000 }).catch(() => {});
                                await page.waitForTimeout(300);
                                console.log(`  거래처명 선택: ${vendorName}`);
                            } else {
                                await page.keyboard.press('Escape');
                                await page.waitForTimeout(300);
                                console.log(`[경고] 거래처명 드롭다운 미발견 — 입력값 그대로 사용: ${vendorName}`);
                            }
                        } else {
                            await page.keyboard.press('Escape');
                            await page.waitForTimeout(200);
                        }
                    } catch { console.log('[경고] 거래처명 combobox를 찾지 못했습니다.'); }

                    // 3. 적요 combobox
                    try {
                        const descCombo = page.locator(COMBO_SEL).nth(2);
                        await descCombo.click();
                        await page.waitForTimeout(200);
                        await page.keyboard.press('Control+A');
                        await page.keyboard.press('Backspace');
                        if (description) {
                            await page.keyboard.type(description, { delay: 50 });
                            await page.waitForTimeout(700);
                            const descOption = page.locator('[role="option"], [role="listbox"] li, ul[role="listbox"] li').first();
                            if (await descOption.count().catch(() => 0) > 0) {
                                await descOption.click({ timeout: 2000 }).catch(() => {});
                                await page.waitForTimeout(300);
                                console.log(`  적요 선택: ${description}`);
                            } else {
                                await page.keyboard.press('Escape');
                                await page.waitForTimeout(300);
                                console.log(`[경고] 적요 드롭다운 미발견 — 입력값 그대로 사용: ${description}`);
                            }
                        } else {
                            await page.keyboard.press('Escape');
                            await page.waitForTimeout(200);
                        }
                    } catch { console.log('[경고] 적요 combobox를 찾지 못했습니다.'); }

                    // 4. 날짜
                    if (startDateRaw) { const d = formatExcelDate(startDateRaw); await page.getByLabel(/시작/i).first().fill(d).catch(() => {}); console.log(`  시작일: ${d}`); }
                    if (endDateRaw)   { const d = formatExcelDate(endDateRaw);   await page.getByLabel(/종료/i).first().fill(d).catch(() => {}); console.log(`  종료일: ${d}`); }

                    // 5. 금액 유형 / 표시 방식 라디오
                    if (amountType)  await clickRadioByLabel(page, amountType,  '금액 유형');
                    if (displayType) await clickRadioByLabel(page, displayType, '표시 방식');

                    // 6. 검색
                    await page.waitForSelector(SEARCH_SEL, { state: 'visible', timeout: 10000 });
                    await page.click(SEARCH_SEL);
                    await page.waitForTimeout(1500);

                    // 7. 결과 다운로드 → groupBook 에 시트 추가
                    const dlVisible = await page.locator(DOWNLOAD_SEL)
                        .waitFor({ state: 'visible', timeout: 8000 })
                        .then(() => true).catch(() => false);
                    if (!dlVisible) {
                        console.log(`  [안내] '${accountName}' 검색 결과 없음 — 다운로드를 건너뜁니다.`);
                    } else {
                        await downloadAndAddSheet(page, DOWNLOAD_SEL, accountName, groupBook);
                    }

                } catch (e) {
                    console.log(`[경고] [${taskName} / ${accountName}] 오류 (다음 계정으로): ${e.message}`);
                }

                // 8. 초기화 (마지막 계정 제외)
                const isLast = gi === allGroups.length - 1 && ti === groupTasks.length - 1;
                if (!isLast) {
                    try {
                        await page.waitForSelector(RESET_SEL, { state: 'visible', timeout: 5000 });
                        await page.click(RESET_SEL);
                        await page.waitForTimeout(1000);
                        await page.waitForSelector(COMBO_SEL, { state: 'visible', timeout: 10000 });
                        console.log('  → [초기화] 완료');
                    } catch (e) {
                        console.log(`[경고] 초기화 실패 (계속 진행): ${e.message}`);
                    }
                }
            }

            // 그룹 파일 저장 (OneDrive EBUSY 재시도)
            for (let attempt = 1; attempt <= 5; attempt++) {
                try {
                    await groupBook.xlsx.writeFile(groupFilePath);
                    console.log(`\n[저장] ${path.relative(ROOT_DIR, groupFilePath)}`);
                    break;
                } catch (e) {
                    if (e.code === 'EBUSY' && attempt < 5) {
                        console.log(`[저장] 잠금 감지, ${attempt}초 후 재시도...`);
                        await new Promise(r => setTimeout(r, attempt * 1000));
                    } else throw e;
                }
            }
        }

    } finally {
        await browser.close();
    }

    console.log('\n=== 완료 ===');
}

main().catch(err => { console.error('[오류]', err.message); process.exit(1); });
