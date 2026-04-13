'use strict';
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });

module.exports = {
    companyName: 'Braintree',

    // 웹 앱 기본 URL (task_list의 Settings 시트에 Url이 있으면 그쪽이 우선)
    url: 'http://127.0.0.1:8081',

    // ── 세션 유지 설정 ─────────────────────────────────────────────────────────
    // 로그인 쿠키를 이 경로에 저장하여 재실행 시 자동 로그인 건너뜀.
    // 처음 실행 시 로그인이 한 번 필요하고, 이후 세션이 유지됨.
    // null 로 설정하면 매번 새 세션(Incognito)으로 실행.
    persistentProfile: path.join(__dirname, '..', '.browser_profile'),

    // ── 메뉴-엔드포인트 커스텀 매핑 (선택사항) ──────────────────────────────────
    // 기본 매핑: '분개장' 계열 → /ai-analysis, 나머지 → /analysis
    // 아래에서 시트명(메뉴명)별로 경로를 덮어쓸 수 있음.
    menuEndpoints: {
        // '커스텀메뉴': '/ai-analysis',
    },

    // ── 파일 설정 ──────────────────────────────────────────────────────────────
    // 분개장 시트에 '업로드파일' 열이 없을 때 기본으로 사용할 업로드 파일
    uploadFileName: 'raw_data/전계정별원장_브랜트리_25년.xlsx',
    templateFileName: 'template.xlsx',
    dataStartRowIndex: 10,

    // ── 인증 정보 ──────────────────────────────────────────────────────────────
    credentials: {
        userId: process.env.USER_EMAIL,
        userPassword: process.env.USER_PASSWORD,
    },

    // ── CSS 셀렉터 ─────────────────────────────────────────────────────────────
    selectors: {
        // /analysis 로그인
        loginId: 'input[type="email"]',
        loginPassword: 'input[type="password"]',
        loginButton: 'button:has-text("로그인"), #login-btn',

        // /analysis 메뉴 공통
        accountCombobox: 'button[role="combobox"]',
        resetButton: 'button:has-text("초기화")',
        searchButton: 'button:has-text("검색")',
        excelDownloadBtn: 'button:has-text("결과 다운로드")',
        resultTable: 'table',

        // /ai-analysis 업로드
        // uploadButton: 'button:has-text("파일 선택")',  // 보이는 업로드 버튼이 있는 경우 활성화
        fileUploadInput: 'input[type="file"]',          // 숨겨진 file input 직접 주입 방식
    },

    // ── 기타 ───────────────────────────────────────────────────────────────────
    targetAccounts: ['미지급금', '외상매입금', '선급금'],
    outputFileName: 'Braintree_Audit_Result',
};
