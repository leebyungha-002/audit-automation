'use strict';
const path = require('path');
const fs   = require('fs');
const ExcelJS = require('exceljs');

// ─── raw_data 폴더에서 파일 자동 감지 ──────────────────────────────────────────
// 당기 계정별원장: "원장" 포함 & "분개장"·"전기" 미포함
// 분개장:         "분개장" 포함
function autoDetectRawFiles(companyDir) {
    const rawDir = path.join(companyDir, 'raw_data');
    if (!fs.existsSync(rawDir)) return {};

    const files = fs.readdirSync(rawDir).filter(f => {
        const ext = path.extname(f).toLowerCase();
        return (ext === '.xlsx' || ext === '.xls') && !f.startsWith('~$');
    });

    const ledger  = files.find(f => /원장/.test(f) && !/분개장/.test(f) && !/전기/.test(f));
    const journal = files.find(f => /분개장/.test(f));

    return {
        ledger:  ledger  ? `raw_data/${ledger}`  : null,
        journal: journal ? `raw_data/${journal}` : null,
    };
}

async function main() {
    console.log('=== 다중 회사 감사 자동화 통합 실행 ===');
    const args = process.argv.slice(2);

    if (args.length === 0) {
        console.error('사용법: node run.js <폴더명>');
        console.error('예시: node run.js Braintree');
        process.exit(1);
    }

    const targetCompany = args[0];
    const companyDir = path.join(__dirname, targetCompany);

    if (!fs.existsSync(companyDir)) {
        console.error(`[오류] 대상 폴더를 찾을 수 없습니다: ${companyDir}`);
        process.exit(1);
    }

    // ── 1. 루트 공통 설정 로드 ───────────────────────────────────────────────
    const sharedConfigPath = path.join(__dirname, 'shared_config.js');
    let config = fs.existsSync(sharedConfigPath) ? { ...require(sharedConfigPath) } : {};

    // ── 2. 회사별 config.js가 있으면 선택적 병합 (없어도 동작) ───────────────
    const companyConfigPath = path.join(companyDir, 'config.js');
    if (fs.existsSync(companyConfigPath)) {
        const override = require(companyConfigPath);
        if (Object.keys(override).length > 0) {
            config = {
                ...config,
                ...override,
                selectors:   { ...config.selectors,   ...(override.selectors   || {}) },
                credentials: { ...config.credentials, ...(override.credentials || {}) },
            };
        }
    }

    // ── 3. 회사명은 항상 폴더명 기준 ────────────────────────────────────────
    config.companyName = targetCompany;

    try {
        // ── 4. task_list 자동 검색 및 읽기 ───────────────────────────────────
        const dirFiles = fs.readdirSync(companyDir);
        const taskListFile =
            dirFiles.find(f => {
                const lf = f.toLowerCase();
                return lf.startsWith('task_list_') && lf.endsWith('.xlsx')
                    && lf.includes(targetCompany.toLowerCase());
            }) ||
            dirFiles.find(f =>
                f.toLowerCase().startsWith('task_list_') && f.toLowerCase().endsWith('.xlsx')
            );

        if (taskListFile) {
            const taskListPath = path.join(companyDir, taskListFile);
            console.log(`[안내] 지시서 파일을 찾았습니다: ${taskListFile}`);

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(taskListPath);
            const taskListData = {};
            const menus = [];

            workbook.eachSheet((sheet) => {
                const sheetName = sheet.name;
                if (sheetName.toLowerCase() === 'settings') {
                    sheet.eachRow((row, rowNum) => {
                        if (rowNum > 1) {
                            const key = row.getCell(2).value;
                            const val = row.getCell(3).value;
                            if (key) taskListData[String(key).trim()] = val;
                        }
                    });
                } else {
                    const tasks = [];
                    let headers = [];
                    sheet.eachRow((row, rowNum) => {
                        const values = row.values;
                        if (rowNum === 1) {
                            for (let i = 1; i < values.length; i++) {
                                headers[i] = values[i] ? values[i].toString().trim() : null;
                            }
                        } else {
                            const task = {};
                            for (let i = 1; i < values.length; i++) {
                                if (headers[i] && values[i] !== undefined && values[i] !== null) {
                                    task[headers[i]] = values[i];
                                }
                            }
                            if (Object.keys(task).length > 0) tasks.push(task);
                        }
                    });
                    menus.push({ menuName: sheetName, tasks });
                }
            });

            console.log(`[안내] 지시서 데이터 로드 완료: settings 항목 ${Object.keys(taskListData).length}개, 메뉴 ${menus.length}개`);
            config.taskList = taskListData;
            config.menus    = menus;
        } else {
            console.log(`[안내] ${targetCompany} 폴더 내에 task_list_*.xlsx 파일이 없습니다.`);
        }

        // ── 5. 업로드 파일명 해결: Settings > company config.js > raw_data 자동감지 ──
        // 우선순위: task_list Settings 항목 → 기존 config.js 값 → raw_data 폴더 자동감지
        const rawDetect = autoDetectRawFiles(companyDir);

        if (!config.uploadFileName) {
            const fromSettings = config.taskList?.LedgerFile ?? config.taskList?.계정별원장파일;
            config.uploadFileName = fromSettings ?? rawDetect.ledger ?? null;
            if (config.uploadFileName)
                console.log(`[안내] 계정별원장 파일 자동 설정: ${config.uploadFileName}`);
        }
        if (!config.aiJournalFileName) {
            const fromSettings = config.taskList?.JournalFile ?? config.taskList?.분개장파일;
            config.aiJournalFileName = fromSettings ?? rawDetect.journal ?? null;
            if (config.aiJournalFileName)
                console.log(`[안내] 분개장 파일 자동 설정: ${config.aiJournalFileName}`);
        }
        if (!config.aiLedgerFileName) {
            config.aiLedgerFileName =
                config.taskList?.AiLedgerFile ?? config.uploadFileName ?? null;
        }

        const runAudit = require('./shared_modules/auditRunner');
        await runAudit(config, companyDir);
    } catch (err) {
        console.error(`[오류] 실행 중 에러가 발생했습니다:`, err);
    }
}

main();
