'use strict';
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });

module.exports = {
    companyName: 'graphy',

    // /analysis 업로드: 당기 계정별원장
    uploadFileName: 'raw_data/당기_graphy_계정별원장_25년.xlsx',

    // /ai-analysis 업로드: 분개장(zone 0) / 계정별원장(zone 1)
    aiJournalFileName: 'raw_data/당기_graphy_분개장_25년.xlsx',
    aiLedgerFileName:  'raw_data/당기_graphy_계정별원장_25년.xlsx',
};
