#!/usr/bin/env node
/* =============================================================================
 * sql_studio_api_bridge.js
 * =============================================================================
 * The HTTP bridge that lets sql_studio_browser_runner_v4.html run real queries
 * against Microsoft SQL Server.
 *
 * WHY THIS FILE EXISTS
 * --------------------
 * A browser cannot connect to SQL Server directly. SQL Server speaks the TDS
 * protocol over a raw TCP socket on port 1433, and browser JavaScript has no raw
 * TCP socket API — only HTTP(S), WebSocket and WebRTC. Integrated Security
 * (SSPI / NTLM) further requires OS-level credential delegation that a sandboxed
 * page is never granted.
 *
 * This process runs on a machine that CAN reach the database. It accepts HTTP
 * POSTs from the HTML page, performs the real TDS connection with the `mssql`
 * driver, and returns rows as JSON.
 *
 *     Browser (HTML page)  --HTTP/JSON-->  this bridge  --TDS/1433-->  SQL Server
 *
 * SETUP
 * -----
 *     npm install express mssql cors
 *     node sql_studio_api_bridge.js
 *
 * Then in the HTML file's APP_CONFIG:
 *     EXECUTION_MODE: "api",
 *     API_ENDPOINT:   "http://localhost:8787/query"
 *
 * For Windows Integrated Security (SSPI), also install the MSAL/NTLM driver:
 *     npm install msnodesqlv8 mssql/msnodesqlv8
 * and set DRIVER = 'msnodesqlv8' below. That path requires the bridge to run on
 * Windows as the account whose credentials should be used.
 *
 * CONTRACT
 * --------
 * REQUEST   POST /query
 *           { sql, host, port, database, auth, requestId }
 *
 * RESPONSE  200  { columns: [...], rows: [...], rowCount, elapsedMs }
 *           400  { error: "message" }
 *
 * SECURITY — READ BEFORE DEPLOYING
 * --------------------------------
 *  * This bridge executes arbitrary SQL. Bind it to localhost (the default) or
 *    put it behind authentication. Do NOT expose it on an open network.
 *  * READ_ONLY below rejects anything that is not a SELECT. Leave it true unless
 *    you have a specific reason and compensating controls.
 *  * Credentials belong in environment variables, never in this file or the HTML.
 *  * Use HTTPS in any shared environment. A browser on an https:// page will
 *    refuse to call an http:// bridge (mixed-content blocking).
 * ============================================================================= */

'use strict';

const express = require('express');
const cors = require('cors');
const sql = require('mssql');

/* ---------------------------------------------------------------------------
 * CONFIGURATION
 * ------------------------------------------------------------------------- */
const CONFIG = {
  // Port this bridge listens on.
  PORT: process.env.BRIDGE_PORT || 8787,

  // Bind address. '127.0.0.1' keeps the bridge unreachable from other machines.
  // Change only if you understand the exposure.
  HOST: process.env.BRIDGE_HOST || '127.0.0.1',

  // Origins permitted to call this bridge. Use '*' only for local development.
  // If you serve the HTML from a web server, list that origin explicitly.
  ALLOWED_ORIGINS: process.env.BRIDGE_ORIGINS
    ? process.env.BRIDGE_ORIGINS.split(',')
    : '*',

  // When true, only SELECT / WITH statements are accepted. Strongly recommended.
  READ_ONLY: process.env.BRIDGE_READ_ONLY !== 'false',

  // Maximum rows returned to the browser, to avoid flooding the page.
  MAX_ROWS: parseInt(process.env.BRIDGE_MAX_ROWS || '5000', 10),

  // Query timeout in milliseconds.
  QUERY_TIMEOUT_MS: parseInt(process.env.BRIDGE_TIMEOUT_MS || '30000', 10),

  // SQL authentication credentials. Supply via environment variables.
  // Ignored when TRUSTED_CONNECTION is true.
  SQL_USER: process.env.SQL_USER || '',
  SQL_PASSWORD: process.env.SQL_PASSWORD || '',

  // Set true to use Windows Integrated Security. Requires msnodesqlv8 and a
  // Windows host running as the intended account.
  TRUSTED_CONNECTION: process.env.SQL_TRUSTED === 'true',

  // Encryption. Azure SQL requires true. On-prem often uses false with
  // trustServerCertificate true for self-signed certs.
  ENCRYPT: process.env.SQL_ENCRYPT === 'true',
  TRUST_SERVER_CERTIFICATE: process.env.SQL_TRUST_CERT !== 'false'
};

/* ---------------------------------------------------------------------------
 * Statement guard
 * ------------------------------------------------------------------------- */
function stripComments(s) {
  return s
    .replace(/\/\*[\s\S]*?\*\//g, ' ')   // block comments
    .replace(/--[^\n]*/g, ' ')           // line comments
    .trim();
}

function assertReadOnly(rawSql) {
  if (!CONFIG.READ_ONLY) return;

  const body = stripComments(rawSql);
  if (!body) throw new Error('Empty statement.');

  // Must begin with SELECT or WITH (CTE).
  if (!/^\s*(SELECT|WITH)\b/i.test(body)) {
    throw new Error(
      'READ_ONLY mode is enabled: only SELECT (or WITH ... SELECT) statements are accepted. ' +
      'Set BRIDGE_READ_ONLY=false to disable, but understand the risk first.'
    );
  }

  // Reject obvious mutation or batch-separation attempts.
  const forbidden = /\b(INSERT|UPDATE|DELETE|DROP|TRUNCATE|ALTER|CREATE|GRANT|REVOKE|EXEC|EXECUTE|MERGE|BACKUP|RESTORE|SHUTDOWN|xp_cmdshell|sp_configure)\b/i;
  const hit = body.match(forbidden);
  if (hit) {
    throw new Error(`READ_ONLY mode is enabled: the statement contains '${hit[1]}', which is not permitted.`);
  }
}

/* ---------------------------------------------------------------------------
 * Connection pooling — one pool per distinct target
 * ------------------------------------------------------------------------- */
const pools = new Map();

async function getPool(host, port, database) {
  const key = `${host}:${port}/${database}`;
  if (pools.has(key)) {
    const existing = pools.get(key);
    if (existing.connected || existing.connecting) return existing;
    pools.delete(key);
  }

  const cfg = {
    server: host,
    port: Number(port) || 1433,
    database,
    connectionTimeout: 15000,
    requestTimeout: CONFIG.QUERY_TIMEOUT_MS,
    pool: { max: 5, min: 0, idleTimeoutMillis: 30000 },
    options: {
      encrypt: CONFIG.ENCRYPT,
      trustServerCertificate: CONFIG.TRUST_SERVER_CERTIFICATE,
      enableArithAbort: true
    }
  };

  if (CONFIG.TRUSTED_CONNECTION) {
    cfg.options.trustedConnection = true;          // requires msnodesqlv8
  } else {
    if (!CONFIG.SQL_USER) {
      throw new Error(
        'No SQL credentials configured. Set SQL_USER and SQL_PASSWORD environment ' +
        'variables, or set SQL_TRUSTED=true for Windows Integrated Security.'
      );
    }
    cfg.user = CONFIG.SQL_USER;
    cfg.password = CONFIG.SQL_PASSWORD;
  }

  const pool = new sql.ConnectionPool(cfg);
  pool.on('error', e => console.error(`[pool ${key}]`, e.message));
  pools.set(key, pool);
  await pool.connect();
  console.log(`[bridge] connected pool -> ${key}`);
  return pool;
}

/* ---------------------------------------------------------------------------
 * HTTP server
 * ------------------------------------------------------------------------- */
const app = express();
app.use(express.json({ limit: '1mb' }));
app.use(cors({ origin: CONFIG.ALLOWED_ORIGINS }));

// Never let a proxy or browser cache a query response.
app.use((req, res, nextMw) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  nextMw();
});

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    readOnly: CONFIG.READ_ONLY,
    pools: [...pools.keys()],
    maxRows: CONFIG.MAX_ROWS,
    time: new Date().toISOString()
  });
});

app.post('/query', async (req, res) => {
  const started = Date.now();
  const { sql: statement, host, port, database, auth, requestId } = req.body || {};

  console.log(`[bridge] ${requestId || '-'} ${host}:${port}/${database} auth="${auth || '-'}"`);

  try {
    if (!statement || !String(statement).trim()) throw new Error('No SQL statement supplied.');
    if (!host) throw new Error('No host supplied.');
    if (!database) throw new Error('No database supplied.');

    assertReadOnly(String(statement));

    const pool = await getPool(host, port, database);
    const result = await pool.request().query(String(statement));

    const recordset = result.recordset || [];
    const truncated = recordset.length > CONFIG.MAX_ROWS;
    const rows = truncated ? recordset.slice(0, CONFIG.MAX_ROWS) : recordset;

    // Column order straight from the driver's metadata, falling back to the
    // first row's keys when metadata is unavailable.
    let columns = [];
    if (result.recordset && result.recordset.columns) {
      columns = Object.keys(result.recordset.columns)
        .sort((a, b) => result.recordset.columns[a].index - result.recordset.columns[b].index);
    } else if (rows.length) {
      columns = Object.keys(rows[0]);
    }

    res.json({
      columns,
      rows,
      rowCount: recordset.length,
      truncated,
      elapsedMs: Date.now() - started
    });
  } catch (err) {
    console.error(`[bridge] error: ${err.message}`);
    res.status(400).json({
      error: err.message,
      elapsedMs: Date.now() - started
    });
  }
});

app.listen(CONFIG.PORT, CONFIG.HOST, () => {
  console.log('='.repeat(72));
  console.log(' SQL Studio API Bridge');
  console.log('='.repeat(72));
  console.log(` Listening   : http://${CONFIG.HOST}:${CONFIG.PORT}`);
  console.log(` Query URL   : http://${CONFIG.HOST}:${CONFIG.PORT}/query`);
  console.log(` Health URL  : http://${CONFIG.HOST}:${CONFIG.PORT}/health`);
  console.log(` Read-only   : ${CONFIG.READ_ONLY ? 'YES (SELECT only)' : 'NO  <-- statements can mutate data'}`);
  console.log(` Auth        : ${CONFIG.TRUSTED_CONNECTION ? 'Windows Integrated (msnodesqlv8)' : (CONFIG.SQL_USER ? `SQL login "${CONFIG.SQL_USER}"` : 'NOT CONFIGURED — set SQL_USER / SQL_PASSWORD')}`);
  console.log(` Max rows    : ${CONFIG.MAX_ROWS}`);
  console.log(` CORS origin : ${CONFIG.ALLOWED_ORIGINS}`);
  console.log('='.repeat(72));
  console.log(' In the HTML APP_CONFIG set:');
  console.log('   EXECUTION_MODE: "api",');
  console.log(`   API_ENDPOINT:   "http://${CONFIG.HOST}:${CONFIG.PORT}/query"`);
  console.log('='.repeat(72));
});

process.on('SIGINT', async () => {
  console.log('\n[bridge] closing pools...');
  for (const [key, pool] of pools) {
    try { await pool.close(); console.log(`[bridge] closed ${key}`); } catch (_) {}
  }
  process.exit(0);
});
