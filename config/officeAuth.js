const redirectUri = process.env.MS_REDIRECT_URI || 'http://localhost:3000/auth/callback';

const officeAuthConfig = {
    tenantId: process.env.MS_TENANT_ID || 'a43b80bf-06e7-47f3-b84f-e718c0d37a71',
    clientId: process.env.MS_CLIENT_ID || '3df0e69d-c01e-4b50-b0e6-81f4712791b9',
    clientSecret: process.env.MS_CLIENT_SECRET || ""
    redirectUri,
    postLogoutRedirectUri: process.env.MS_POST_LOGOUT_REDIRECT_URI || new URL(redirectUri).origin,
    scopes: ['openid', 'profile', 'email']
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

officeAuthConfig.isConfigured = Boolean(
    officeAuthConfig.tenantId &&
    officeAuthConfig.clientId &&
    officeAuthConfig.clientSecret
);

module.exports = officeAuthConfig;