const path = require('path');
const express = require('express');
const session = require('express-session');
const msal = require('@azure/msal-node');
const pool = require('./config/database');
const officeAuthConfig = require('./config/officeAuth');

const app = express();
const port = process.env.PORT || 3000;
const msalClient = new msal.ConfidentialClientApplication(officeAuthConfig.msalConfig);
const cryptoProvider = new msal.CryptoProvider();

app.use(express.json());
app.use(session({
    name: 'ramais.office.sid',
    secret: process.env.SESSION_SECRET || 'troque-esta-chave-de-sessao',
    resave: false,
    saveUninitialized: false,
    cookie: {
        httpOnly: true,
        sameSite: 'lax',
        secure: process.env.NODE_ENV === 'production'
    }
}));

function normalizeGroup(value) {
    return String(value || '')
        .trim()
        .toUpperCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
}

function serializeRamal(row) {
    const grupoNormalizado = normalizeGroup(row.grupo);
    const ramal = row.ramal === null || row.ramal === undefined ? '' : String(row.ramal);
    const item = {
        id: row.id,
        grupo: row.grupo || '',
        nome: row.nome || '',
        setor: row.setor || '',
        ramal
    };

    if (grupoNormalizado === 'BASES' || grupoNormalizado === 'AUTO POSTO') {
        item.local = row.setor || '';
    }

    if (grupoNormalizado === 'COMANDOS') {
        item.descricao = row.setor || '';
    }

    return item;
}

function requireAuth(req, res, next) {
    if (req.session.officeUser) {
        next();
        return;
    }

    res.status(401).json({ error: 'Login Microsoft 365 necessario.' });
}

function getLogoutUrl() {
    const url = new URL(`${officeAuthConfig.authority}/oauth2/v2.0/logout`);
    url.searchParams.set('post_logout_redirect_uri', officeAuthConfig.postLogoutRedirectUri);
    return url.toString();
}

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'ramais.html'));
});

app.get('/favicon.ico', (req, res) => {
    res.sendFile(path.join(__dirname, 'favicon.ico'));
});

app.get('/auth/status', (req, res) => {
    res.json({
        configured: officeAuthConfig.isConfigured,
        authenticated: Boolean(req.session.officeUser),
        user: req.session.officeUser || null,
        authError: req.session.authError || null
    });
});

app.get('/auth/login', async (req, res) => {
    if (!officeAuthConfig.isConfigured) {
        res.status(500).send('Login Microsoft 365 ainda nao configurado. Preencha config/officeAuth.js.');
        return;
    }

    req.session.authError = null;

    try {
        const state = cryptoProvider.createNewGuid();
        const nonce = cryptoProvider.createNewGuid();
        req.session.authState = state;
        req.session.authNonce = nonce;

        const authUrl = await msalClient.getAuthCodeUrl({
            scopes: officeAuthConfig.scopes,
            redirectUri: officeAuthConfig.redirectUri,
            state,
            nonce,
            prompt: 'select_account'
        });

        res.redirect(authUrl);
    } catch (error) {
        console.error('Erro ao iniciar login Microsoft 365:', error);
        res.status(500).send('Nao foi possivel iniciar o login Microsoft 365.');
    }
});

app.get('/auth/callback', async (req, res) => {
    if (req.query.error) {
        console.error('Erro retornado pela Microsoft:', req.query.error, req.query.error_description);
        res.redirect('/');
        return;
    }

    if (!req.query.code || req.query.state !== req.session.authState) {
        res.status(400).send('Resposta de login invalida.');
        return;
    }

    try {
        const tokenResponse = await msalClient.acquireTokenByCode({
            code: req.query.code,
            scopes: officeAuthConfig.scopes,
            redirectUri: officeAuthConfig.redirectUri
        });

        const objectId = tokenResponse.idTokenClaims?.oid || tokenResponse.account?.localAccountId || '';

        req.session.officeUser = {
            name: tokenResponse.account?.name || tokenResponse.idTokenClaims?.name || '',
            username: tokenResponse.account?.username || tokenResponse.idTokenClaims?.preferred_username || '',
            tenantId: tokenResponse.account?.tenantId || tokenResponse.idTokenClaims?.tid || '',
            objectId
        };

        req.session.authError = null;
        delete req.session.authState;
        delete req.session.authNonce;

        res.redirect('/');
    } catch (error) {
        console.error('Erro no callback Microsoft 365:', error);
        res.status(500).send(`
<pre>
${error.stack || error.message || JSON.stringify(error, null, 2)}
</pre>
        `);
    }
});

app.get('/auth/logout', (req, res) => {
    const logoutUrl = getLogoutUrl();
    req.session.destroy(() => {
        res.clearCookie('ramais.office.sid');
        res.redirect(logoutUrl);
    });
});

app.get('/ramais', requireAuth, async (req, res) => {
    try {
        const result = await pool.query(
            `SELECT id, grupo, nome, setor, ramal
               FROM ramais
              ORDER BY id`
        );

        res.json(result.rows.map(serializeRamal));
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(port, () => {
    console.log(`Servidor de ramais rodando em http://localhost:${port}`);
});
