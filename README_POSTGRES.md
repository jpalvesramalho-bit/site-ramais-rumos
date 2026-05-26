# PostgreSQL e DBeaver

## Configuracao

1. Crie um banco no PostgreSQL, por exemplo `ramais`.
2. Abra o DBeaver, conecte nesse banco e execute o arquivo `database/seed_ramais.sql`.
3. Configure a conexao do site em `config/database.js`:

```js
const databaseConfig = {
    host: process.env.PGHOST || 'localhost',
    port: Number(process.env.PGPORT || 5432),
    user: process.env.PGUSER || 'postgres',
    password: process.env.PGPASSWORD || 'SUA_SENHA_POSTGRES',
    database: process.env.PGDATABASE || 'ramais'
};
```

## Executar

```bash
npm start
```

Depois acesse:

```text
http://localhost:3000
```

Ao alterar qualquer registro da tabela `ramais` pelo DBeaver, salve a alteracao e atualize o navegador. O site buscara os dados novamente pelo endpoint `GET /ramais`.
