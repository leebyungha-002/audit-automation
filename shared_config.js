'use strict';
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });

module.exports = {
    // 기본 URL — task_list Settings의 Url 항목이 있으면 그쪽이 우선 적용됨
    url: 'http://127.0.0.1:8081',

    // 세션 유지용 브라우저 프로필 (null이면 매번 새 세션)
    persistentProfile: path.join(__dirname, '.browser_profile'),

    templateFileName: 'template.xlsx',
    dataStartRowIndex: 10,
    menuEndpoints: {},

    credentials: {
        userId: process.env.USER_EMAIL,
        userPassword: process.env.USER_PASSWORD,
    },

    selectors: {
        loginId: 'input[type="email"]',
        loginPassword: 'input[type="password"]',
        loginButton: 'button:has-text("로그인"), #login-btn',
        accountCombobox: 'button[role="combobox"]',
        resetButton: 'button:has-text("초기화")',
        searchButton: 'button:has-text("검색")',
        excelDownloadBtn: 'button:has-text("결과 다운로드")',
        resultTable: 'table',
        fileUploadInput: 'input[type="file"]',
    },
};
