const sso = require('./node_modules/office-addin-sso/lib/server');
require('dotenv').config();

const ssoOptions = {
    applicationId: process.env.CLIENT_ID,
    applicationName: 'Office Add-in NodeJS SSO',
    graphApiScopes: ['Files.Read.All'],
    queryParam: '?$select=name&$top=5',
    tenantId: process.env.TENANT_ID
}

const ssoInstance = new sso.SSOService(ssoOptions, process.env.PORT);
ssoInstance.startSsoService();