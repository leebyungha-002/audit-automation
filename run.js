const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

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
    
    // 타겟 폴더 존재 여부 확인
    if (!fs.existsSync(companyDir)) {
        console.error(`[오류] 대상 폴더를 찾을 수 없습니다: ${companyDir}`);
        process.exit(1);
    }
    
    const configPath = path.join(companyDir, 'config.js');
    
    // 설정 파일 존재 여부 확인
    if (!fs.existsSync(configPath)) {
        console.error(`[오류] 해당 폴더에 config.js가 없습니다: ${configPath}`);
        console.error('공통 로직을 사용하려면 각 회사 폴더에 config.js가 반드시 필요합니다.');
        process.exit(1);
    }
    
    try {
        const config = require(configPath);
        
        // 지시서(task_list_*.xlsx) 자동 검색 및 읽기
        const files = fs.readdirSync(companyDir);
        const taskListFile = files.find(f => {
            const lowerF = f.toLowerCase();
            const lowerCompany = targetCompany.toLowerCase();
            return lowerF.startsWith('task_list_') && lowerF.endsWith('.xlsx') && lowerF.includes(lowerCompany);
        }) || files.find(f => f.toLowerCase().startsWith('task_list_') && f.toLowerCase().endsWith('.xlsx'));
        
        if (taskListFile) {
            const taskListPath = path.join(companyDir, taskListFile);
            console.log(`[안내] 지시서 파일을 찾았습니다: ${taskListFile}`);
            
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(taskListPath);
            const taskListData = {};
            const menus = [];
            
            workbook.eachSheet((sheet) => {
                const sheetName = sheet.name;
                
                if (sheetName.toLowerCase() === 'settings' || sheetName === 'Settings') {
                    sheet.eachRow((row, rowNum) => {
                        if (rowNum > 1) { // 1행(헤더) 구조를 감안하여 2행부터 데이터 읽기
                            // B열(2)이 '항목', C열(3)이 '설정값' 가정
                            const key = row.getCell(2).value;
                            const val = row.getCell(3).value;
                            if (key) {
                                taskListData[key] = val;
                            }
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
                            if (Object.keys(task).length > 0) {
                                tasks.push(task);
                            }
                        }
                    });
                    menus.push({ menuName: sheetName, tasks });
                }
            });
            
            console.log(`[안내] 지시서 데이터 로드 완료: settings 항목 ${Object.keys(taskListData).length}개, 메뉴 ${menus.length}개`);
            
            // 기존 config에 지시서 내용 병합
            config.taskList = taskListData;
            config.menus = menus;
        } else {
            console.log(`[안내] ${targetCompany} 폴더 내에 task_list_*.xlsx 지시서 파일이 없습니다.`);
        }

        const runAudit = require('./shared_modules/auditRunner');
        
        await runAudit(config, companyDir);
    } catch (err) {
        console.error(`[오류] 실행 중 에러가 발생했습니다:`, err);
    }
}

main();
