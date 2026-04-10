module.exports = {
    url: 'https://example-company-b.com/admin/login',
    selectors: {
        loginId: '#userId',
        loginPassword: '#userPwd',
        loginButton: '.btn-login',
        auditDataRow: 'tr.data-item'
    },
    outputFileName: 'CompanyB_Audit_Result.xlsx'
};
