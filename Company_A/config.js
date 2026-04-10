const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });

module.exports = {
    companyName: 'Company_A',
    url: 'http://127.0.0.1:8081/', // Company_A 접속 주소 교체 필요 시 변경
    uploadFileName: 'raw_data/Company_A_원장.xlsx', // Company_A용 업로드 파일 (경로 생성 필요)
    templateFileName: 'template.xlsx',
    dataStartRowIndex: 2, // Company_A의 경우 헤더가 1줄이면 2번째 줄부터 (필요 시 수정)
    credentials: {
        userId: process.env.USER_EMAIL,
        userPassword: process.env.USER_PASSWORD
    },
    selectors: {
        loginId: 'input[type="email"]',
        loginPassword: 'input[type="password"]',
        loginButton: 'button:has-text("로그인"), #login-btn',
        analysisMenu: 'text="상세 거래 검색"',
        excelDownloadBtn: 'button:has-text("결과 다운로드")',
        fileUploadInput: 'input[type="file"]',
        accountCombobox: 'button[role="combobox"]',
        resetButton: 'button:has-text("초기화")',
        searchButton: 'button:has-text("검색")',
        resultTable: 'table'
    },
    targetAccounts: ['미지급금'], // Company_A의 대상 계정과목
    outputFileName: 'Company_A_Audit_Result'
};
