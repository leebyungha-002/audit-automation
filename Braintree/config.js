const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });

module.exports = {
    companyName: 'Braintree',
    url: 'http://127.0.0.1:8081/', // 업로드 화면으로 접속
    uploadFileName: 'raw_data/전계정별원장_브랜트리_25년.xlsx',
    templateFileName: 'template.xlsx',
    dataStartRowIndex: 10,
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
    targetAccounts: ['미지급금', '외상매입금', '선급금'], 
    outputFileName: 'Braintree_Audit_Result' // 확장명 생략 시 알아서 추가됨
};
