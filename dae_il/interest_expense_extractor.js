'use strict';
const path    = require('path');
const fs      = require('fs');
const ExcelJS = require('exceljs');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const { initBrowser } = require('../shared_modules/index');

const COMPANY_DIR    = __dirname;
const TASK_LIST_PATH = path.join(COMPANY_DIR, 'task_list_dae_il.xlsx');
const RESULTS_DIR    = path.join(COMPANY_DIR, 'results');
const PROFILE_DIR    = path.join(__dirname, '..', '.browser_profile');
const BASE_URL       = process.env.AUDIT_URL || 'http://127.0.0.1:8081';

// 탐색 순서대로 시트명 후보 (첫 번째 매칭 사용)
const SHEET_PATTERNS = [
    'host1_이자비용적정성test',
    '이자비용적정성test',
    'host1_이자비용test',
    '이자비용test',
];

const SEL = {
    accountCombobox: 'button[role="combobox"]',
    resetButton:     'button:has-text("초기화")',
    searchButton:    'button:has-text("검색")',
};

// ─── 시트 로드 ────────────────────────────────────────────────────────────────
async function loadTasks() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TASK_LIST_PATH);

    let sheet = null;
    for (const name of SHEET_PATTERNS) {
        sheet = wb.getWorksheet(name);
        if (sheet) break;
    }
    if (!sheet) {
        throw new Error(`이자비용적정성test 시트를 찾을 수 없습니다. (탐색 목록: ${SHEET_PATTERNS.join(', ')})`);
    }

    let headers = null;
    const tasks = [];

    sheet.eachRow((row, rowNum) => {
        const vals = row.values; // ExcelJS: vals[0]은 항상 undefined, 데이터는 vals[1]~
        if (rowNum === 1) {
            headers = vals.slice(1).map(h => (h != null ? String(h).trim() : null));
            return;
        }
        if (!headers) return;

        const task = {};
        headers.forEach((h, i) => {
            if (h) task[h] = vals[i + 1] ?? null;
        });

        const account = String(task['계정과목'] ?? '').trim();
        if (account) tasks.push(task);
    });

    console.log(`[시트] '${sheet.name}' → ${tasks.length}개 행 로드 완료`);
    return tasks;
}

// ─── 라디오 버튼 클릭 헬퍼 ───────────────────────────────────────────────────
async function clickRadioByLabel(page, labelText) {
    if (!labelText) return;
    const text = String(labelText).trim();

    // 전략 1: 텍스트를 직접 포함하는 레이블/역할 요소
    const directSelectors = [
        `label:has-text("${text}")`,
        `div[role="radio"]:has-text("${text}")`,
        `span:has-text("${text}")`,
        `button:has-text("${text}")`,
    ];
    for (const sel of directSelectors) {
        try {
            const loc = page.locator(sel).first();
            if (await loc.count() > 0) {
                await loc.click({ timeout: 3000 });
                await page.waitForTimeout(300);
                console.log(`  라디오 선택: "${text}"`);
                return;
            }
        } catch { /* 다음 전략 */ }
    }

    // 전략 2: input[type="radio"] 의 id 와 연결된 label 텍스트 매칭
    try {
        const inputs = page.locator('input[type="radio"]');
        const count  = await inputs.count();
        for (let i = 0; i < count; i++) {
            const inp = inputs.nth(i);
            const id  = await inp.getAttribute('id').catch(() => null);
            if (id) {
                const lbl    = page.locator(`label[for="${id}"]`);
                const lblTxt = (await lbl.innerText().catch(() => '')).trim();
                if (lblTxt.includes(text)) {
                    await inp.click({ timeout: 3000 });
                    await page.waitForTimeout(300);
                    console.log(`  라디오 선택(id 매핑): "${text}"`);
                    return;
                }
            }
        }
    } catch { /* 실패 */ }

    console.log(`[경고] '${text}' 라디오 버튼을 찾지 못했습니다.`);
}

// ─── 결과 테이블 파싱 ─────────────────────────────────────────────────────────
// 검색 후 화면에 나타난 결과 테이블을 DOM에서 직접 파싱.
// 가장 행이 많은 <table> 을 결과 테이블로 간주.
// thead/tbody 구조를 우선하고, 없으면 첫 번째 <tr> 을 헤더로 사용.
async function parseResultTable(page) {
    // UI 상의 "결과 없음" 메시지 확인
    const noResultCount = await page.locator(
        'text="검색 결과가 없습니다", text="결과가 없습니다", text="데이터가 없습니다", text="조회된 데이터가 없습니다"'
    ).count().catch(() => 0);
    if (noResultCount > 0) {
        return [];
    }

    // 테이블 출현 대기 (최대 5초)
    const tableVisible = await page.waitForSelector('table', { state: 'visible', timeout: 5000 })
        .then(() => true)
        .catch(() => false);
    if (!tableVisible) return [];

    return await page.evaluate(() => {
        const tables = document.querySelectorAll('table');
        if (!tables.length) return [];

        // 행이 가장 많은 테이블 = 결과 테이블
        const table = [...tables].reduce((a, b) =>
            a.querySelectorAll('tr').length >= b.querySelectorAll('tr').length ? a : b
        );

        let headerEls = [];
        let dataRows  = [];

        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');

        if (thead && tbody) {
            const hRow = thead.querySelector('tr');
            headerEls  = hRow ? Array.from(hRow.querySelectorAll('th, td')) : [];
            dataRows   = Array.from(tbody.querySelectorAll('tr'));
        } else {
            const allRows = Array.from(table.querySelectorAll('tr'));
            if (!allRows.length) return [];
            headerEls = Array.from(allRows[0].querySelectorAll('th, td'));
            dataRows  = allRows.slice(1);
        }

        const headers = headerEls.map(el => el.textContent.trim());
        if (!headers.length) return [];

        return dataRows.map(tr => {
            const cells = Array.from(tr.querySelectorAll('td, th'));
            const obj   = {};
            cells.forEach((cell, i) => {
                const key = headers[i] || `col${i}`;
                obj[key]  = cell.textContent.trim();
            });
            return obj;
        }).filter(row => Object.values(row).some(v => v !== ''));
    });
}

// ─── CSV 변환 (BOM 포함, Excel 한국어 호환) ──────────────────────────────────
function toCSV(records) {
    if (!records.length) return '';
    const headers  = Object.keys(records[0]);
    const escape   = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
    const dataRows = records.map(r => headers.map(h => escape(r[h])).join(','));
    return '﻿' + [headers.map(escape).join(','), ...dataRows].join('\r\n');
}

// ─── OneDrive EBUSY 재시도 파일 쓰기 ─────────────────────────────────────────
async function writeFileWithRetry(filePath, content) {
    for (let attempt = 1; attempt <= 5; attempt++) {
        try {
            fs.writeFileSync(filePath, content, 'utf8');
            return;
        } catch (e) {
            if (e.code === 'EBUSY' && attempt < 5) {
                console.log(`[파일저장] 잠금 감지, ${attempt}초 후 재시도... (${path.basename(filePath)})`);
                await new Promise(r => setTimeout(r, attempt * 1000));
            } else {
                throw e;
            }
        }
    }
}

// ─── 계정별원장 파일 자동 탐색 ───────────────────────────────────────────────
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
    console.log('=== 이자비용 적정성 데이터 추출 시작 ===');
    console.log(`대상 파일: ${TASK_LIST_PATH}`);

    // 1. 작업 목록 로드
    const tasks = await loadTasks();
    if (!tasks.length) {
        console.log('[종료] 처리할 작업이 없습니다.');
        return;
    }

    // 2. 결과 디렉토리 보장
    if (!fs.existsSync(RESULTS_DIR)) fs.mkdirSync(RESULTS_DIR, { recursive: true });

    // 3. 브라우저 시작 (세션 유지 모드, headless:false 로 동작 확인)
    const { browser, page } = await initBrowser(false, PROFILE_DIR);
    const allResults = [];

    try {
        // 4. /analysis 접속
        const targetUrl = `${BASE_URL}/analysis`;
        console.log(`\n[브라우저] ${targetUrl} 접속 중...`);
        await page.goto(targetUrl, { waitUntil: 'networkidle', timeout: 30000 });
        await page.waitForTimeout(1500);

        // 5. 로그인 화면이면 자격증명 입력
        const isLoginPage = await page.locator('input[type="password"]').count().catch(() => 0) > 0;
        if (isLoginPage) {
            const userId = process.env.USER_EMAIL    ?? '';
            const userPw = process.env.USER_PASSWORD ?? '';
            if (!userId || !userPw) throw new Error('.env 에 USER_EMAIL / USER_PASSWORD 가 없습니다.');

            console.log('[로그인] 자격증명 입력...');
            await page.fill('input[type="email"]',    userId);
            await page.fill('input[type="password"]', userPw);
            await page.click('button:has-text("로그인"), #login-btn');
            await page.waitForLoadState('networkidle', { timeout: 15000 });
            await page.waitForTimeout(1500);
            console.log('[로그인] 완료');
        }

        // 6. 파일 업로드 영역이 있으면 계정별원장 업로드
        const hasFileInput = await page.locator('input[type="file"]').count().catch(() => 0) > 0;
        if (hasFileInput) {
            const ledgerFile = findLedgerFile();
            if (ledgerFile) {
                console.log(`[업로드] 계정별원장: ${path.basename(ledgerFile)}`);
                await page.setInputFiles('input[type="file"]', ledgerFile);

                // 업로드 완료 대기 (파일 input 이 사라지거나 5초 후)
                try {
                    await page.waitForSelector('input[type="file"]', { state: 'hidden', timeout: 30000 });
                } catch {
                    await page.waitForTimeout(5000);
                }

                // 전기 데이터 다이얼로그: "아니요" 클릭 (없으면 무시)
                try {
                    const skipBtn = await page.waitForSelector('button:has-text("아니요")', { state: 'visible', timeout: 5000 });
                    await skipBtn.click();
                    await page.waitForTimeout(1000);
                } catch { /* 다이얼로그 없음 */ }

                console.log('[업로드] 완료');
            } else {
                console.log('[업로드] raw_data 에 계정별원장 파일이 없습니다. 업로드를 건너뜁니다.');
            }
        }

        // 7. '상세 거래 검색' 카드 클릭
        const CARD_LABEL = '상세 거래 검색';
        console.log(`\n[카드] '${CARD_LABEL}' 진입 시도...`);
        try {
            const card = page.locator(
                `text="${CARD_LABEL}", h2:has-text("${CARD_LABEL}"), ` +
                `h3:has-text("${CARD_LABEL}"), div:has-text("${CARD_LABEL}")`
            ).first();
            await card.waitFor({ state: 'visible', timeout: 10000 });
            // 새 탭 방지
            await card.evaluate(n => {
                n.removeAttribute('target');
                n.closest('a')?.removeAttribute('target');
            });
            await card.click();
            await page.waitForTimeout(2000);
            console.log(`[카드] '${CARD_LABEL}' 진입 완료`);
        } catch (e) {
            console.log(`[경고] 카드 클릭 실패 (이미 해당 화면일 수 있음): ${e.message}`);
        }

        // 8. 검색 폼(combobox) 준비 확인
        await page.waitForSelector(SEL.accountCombobox, { state: 'visible', timeout: 15000 });
        console.log('[검색폼] 준비 완료\n');

        // 9. 태스크 반복 처리
        for (let i = 0; i < tasks.length; i++) {
            const task        = tasks[i];
            const accountName = String(task['계정과목'] ?? '').trim();
            const amountType  = String(task['금액 유형'] ?? task['금액유형'] ?? '').trim();
            const displayType = String(task['표시방식']  ?? task['표시 방식']  ?? '').trim();

            if (!accountName) continue;
            console.log(`--- [${i + 1}/${tasks.length}] ${accountName} ---`);

            try {
                // 9-1. 계정과목 combobox 입력
                await page.waitForSelector(SEL.accountCombobox, { state: 'visible', timeout: 10000 });
                await page.waitForTimeout(300);
                await page.click(SEL.accountCombobox);
                await page.waitForTimeout(300);
                await page.keyboard.press('Control+A');
                await page.keyboard.press('Backspace');
                await page.keyboard.type(accountName, { delay: 50 });
                await page.waitForTimeout(500);
                await page.keyboard.press('Enter');
                await page.waitForTimeout(400);
                console.log(`  계정과목: ${accountName}`);

                // 9-2. 금액 유형 라디오
                if (amountType)  await clickRadioByLabel(page, amountType);

                // 9-3. 표시 방식 라디오
                if (displayType) await clickRadioByLabel(page, displayType);

                // 9-4. 검색
                await page.waitForSelector(SEL.searchButton, { state: 'visible', timeout: 10000 });
                await page.click(SEL.searchButton);
                await page.waitForTimeout(1500);

                // 9-5. 결과 테이블 파싱
                const rows = await parseResultTable(page);
                console.log(`  결과: ${rows.length}건`);

                // 9-6. 계정과목 컬럼 주입 후 누적
                allResults.push(...rows.map(row => ({ 계정과목: accountName, ...row })));

            } catch (e) {
                console.log(`[경고] [${accountName}] 처리 중 오류 (다음으로 진행): ${e.message}`);
            }

            // 9-7. 다음 태스크를 위한 초기화 (마지막 제외)
            if (i < tasks.length - 1) {
                try {
                    await page.waitForSelector(SEL.resetButton, { state: 'visible', timeout: 5000 });
                    await page.click(SEL.resetButton);
                    await page.waitForTimeout(1000);
                    await page.waitForSelector(SEL.accountCombobox, { state: 'visible', timeout: 10000 });
                    console.log('  → 초기화 완료');
                } catch (e) {
                    console.log(`[경고] 초기화 실패 (계속 진행): ${e.message}`);
                }
            }
        }

    } finally {
        await browser.close();
    }

    // 10. 결과 저장
    console.log(`\n=== 총 ${allResults.length}건 추출 완료 — 저장 중... ===`);

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const jsonPath  = path.join(RESULTS_DIR, `이자비용적정성_${timestamp}.json`);
    const csvPath   = path.join(RESULTS_DIR, `이자비용적정성_${timestamp}.csv`);

    await writeFileWithRetry(jsonPath, JSON.stringify(allResults, null, 2));
    console.log(`[저장] JSON: ${path.basename(jsonPath)}`);

    if (allResults.length > 0) {
        await writeFileWithRetry(csvPath, toCSV(allResults));
        console.log(`[저장] CSV : ${path.basename(csvPath)}`);
    }

    console.log('\n=== 완료 ===');
}

main().catch(err => {
    console.error('[오류]', err.message);
    process.exit(1);
});
