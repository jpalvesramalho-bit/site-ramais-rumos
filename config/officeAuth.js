const redirectUri = process.env.MS_REDIRECT_URI || 'http://localhost:3000/auth/callback';

const officeAuthConfig = {
    tenantId: process.env.MS_TENANT_ID || 'COLOQUE_SEU_TENANT_ID',
    clientId: process.env.MS_CLIENT_ID || 'COLOQUE_SEU_CLIENT_ID',
    clientSecret: process.env.MS_CLIENT_SECRET || 'COLOQUE_SEU_CLIENT_SECRET',
    redirectUri,
    postLogoutRedirectUri: process.env.MS_POST_LOGOUT_REDIRECT_URI || new URL(redirectUri).origin,
    scopes: ['openid', 'profile', 'email'],
    graphScopes: ['https://graph.microsoft.com/.default'],
    requiredSkuIds: (process.env.MS_REQUIRED_SKU_IDS || '')
        .split(',')
        .map((value) => value.trim().toLowerCase())
        .filter(Boolean)
};

officeAuthConfig.authority = `https://login.microsoftonline.com/${officeAuthConfig.tenantId}`;

officeAuthConfig.msalConfig = {
    auth: {
        clientId: officeAuthConfig.clientId,
        authority: officeAuthConfig.authority,
        clientSecret: officeAuthConfig.clientSecret
    },
    system: {
        loggerOptions: {
            piiLoggingEnabled: false
        }
    }
};

officeAuthConfig.isConfigured = ![
    officeAuthConfig.tenantId,
    officeAuthConfig.clientId,
    officeAuthConfig.clientSecret
].some((value) => value.startsWith('COLOQUE_SEU_'));

module.exports = officeAuthConfig;
