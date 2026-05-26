# Login Microsoft 365

O projeto usa o pacote oficial `@azure/msal-node` para login com Microsoft 365 / Microsoft Entra ID.

## 1. Criar o App Registration

No Microsoft Entra admin center:

1. Entre em `Identity > Applications > App registrations`.
2. Crie um novo registro.
3. Em `Supported account types`, escolha contas desta organizacao.
4. Em `Redirect URI`, selecione `Web` e informe:

```text
http://localhost:3000/auth/callback
```

## 2. Criar o Client Secret

No App Registration:

1. Abra `Certificates & secrets`.
2. Crie um novo `Client secret`.
3. Copie o campo `Value` imediatamente.

## 3. Permitir leitura de licencas

Para liberar apenas usuarios com licenca Microsoft 365, o servidor consulta o Microsoft Graph depois do login.

No App Registration:

1. Abra `API permissions`.
2. Clique em `Add a permission`.
3. Escolha `Microsoft Graph`.
4. Escolha `Application permissions`.
5. Adicione `User.Read.All`.
6. Clique em `Grant admin consent`.

O usuario comum nao precisa ser administrador. A permissao e concedida ao aplicativo uma vez pelo administrador.

## 4. Configurar o projeto

Preencha estes campos em `config/officeAuth.js`:

```js
tenantId: process.env.MS_TENANT_ID || 'COLOQUE_SEU_TENANT_ID',
clientId: process.env.MS_CLIENT_ID || 'COLOQUE_SEU_CLIENT_ID',
clientSecret: process.env.MS_CLIENT_SECRET || 'COLOQUE_SEU_CLIENT_SECRET',
```

Por padrao, qualquer usuario ativo com pelo menos uma licenca atribuida entra. Para exigir uma licenca especifica, preencha `MS_REQUIRED_SKU_IDS` com um ou mais `skuId`, separados por virgula:

```powershell
$env:MS_REQUIRED_SKU_IDS="sku-id-da-licenca"
```

Ou configure por variaveis de ambiente:

```powershell
$env:MS_TENANT_ID="seu-tenant-id"
$env:MS_CLIENT_ID="seu-client-id"
$env:MS_CLIENT_SECRET="seu-client-secret"
$env:SESSION_SECRET="uma-chave-grande-e-aleatoria"
npm start
```

## 5. Usar

Abra:

```text
http://localhost:3000
```

O site mostra o botao `Entrar com Microsoft 365`. Depois do login, o endpoint `GET /ramais` fica liberado para a sessao autenticada.

Se a conta nao tiver licenca atribuida, o login Microsoft acontece, mas o site nao libera os ramais.
