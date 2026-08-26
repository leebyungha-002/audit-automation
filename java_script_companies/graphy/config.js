'use strict';
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../../.env') });

module.exports = {
    companyName: 'graphy',

    // 업로드 파일명: raw_data/current/ 폴더에서 자동 감지 (별도 지정 불필요)
};
