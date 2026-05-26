const { Pool } = require('pg');

// Configure aqui os mesmos dados da conexao usada no DBeaver.
// Tambem e possivel configurar por variaveis de ambiente:
// PGHOST, PGPORT, PGUSER, PGPASSWORD e PGDATABASE.
const databaseConfig = {
    host: process.env.PGHOST || 'localhost',
    port: Number(process.env.PGPORT || 5432),
    user: process.env.PGUSER || 'postgres',
    password: process.env.PGPASSWORD || 'Joao*2026#!',
    database: process.env.PGDATABASE || 'ramais'
};

module.exports = new Pool(databaseConfig);
module.exports.databaseConfig = databaseConfig;
