const sql = require('mssql');

const rawServer = (process.env.DB_SERVER || 'localhost').trim();
const isLocalDB = rawServer.toLowerCase().includes('(localdb)');

const DB_CONFIG = {
  server: isLocalDB ? 'localhost' : rawServer,
  database: process.env.DB_NAME || 'WordToHTMLX',
  options: {
    encrypt: process.env.DB_ENCRYPT === 'true',
    trustServerCertificate: process.env.DB_TRUST_SERVER !== 'false',
    connectTimeout: parseInt(process.env.DB_CONNECT_TIMEOUT || '30000', 10),
    requestTimeout: parseInt(process.env.DB_REQUEST_TIMEOUT || '30000', 10),
  },
  pool: {
    max: parseInt(process.env.DB_POOL_MAX || '10', 10),
    min: parseInt(process.env.DB_POOL_MIN || '0', 10),
    idleTimeoutMillis: parseInt(process.env.DB_IDLE_TIMEOUT || '30000', 10),
  },
};

// LocalDB uses named pipes via instance name
if (isLocalDB) {
  DB_CONFIG.options.instanceName = 'MSSQLLocalDB';
  if (process.env.DB_USER) {
    DB_CONFIG.user = process.env.DB_USER;
    DB_CONFIG.password = process.env.DB_PASSWORD || '';
  } else {
    DB_CONFIG.options.trustedConnection = true;
  }
} else {
  DB_CONFIG.port = parseInt(process.env.DB_PORT || '1433', 10);
  if (process.env.DB_USER) DB_CONFIG.user = process.env.DB_USER;
  if (process.env.DB_PASSWORD) DB_CONFIG.password = process.env.DB_PASSWORD;
}

let pool = null;

async function getPool() {
  if (pool && pool.connected) return pool;
  const needsAuth = !DB_CONFIG.options.trustedConnection && !DB_CONFIG.password;
  if (needsAuth) {
    return null;
  }
  pool = await sql.connect(DB_CONFIG);
  return pool;
}

async function upsertSectionContent({ sectionTitle, sectionNumber, htmlContent, canvasContent, apiResponse, navigationConfig }) {
  const p = await getPool();
  if (!p) return null;

  const existing = await p.request()
    .input('sectionTitle', sql.NVarChar(500), sectionTitle)
    .query('SELECT SectionContentID FROM SectionContent WHERE SectionTitle = @sectionTitle');

  if (existing.recordset.length > 0) {
    const result = await p.request()
      .input('sectionTitle', sql.NVarChar(500), sectionTitle)
      .input('sectionNumber', sql.Int, sectionNumber)
      .input('htmlContent', sql.NVarChar(sql.MAX), htmlContent)
      .input('canvasContent', sql.NVarChar(sql.MAX), canvasContent)
      .input('apiResponse', sql.NVarChar(sql.MAX), apiResponse)
      .input('navigationConfig', sql.NVarChar(sql.MAX), navigationConfig)
      .query(`UPDATE SectionContent
              SET SectionNumber = @sectionNumber,
                  HtmlContent = @htmlContent,
                  CanvasContent = @canvasContent,
                  APIResponse = @apiResponse,
                  NavigationConfig = @navigationConfig,
                  UpdatedAt = GETUTCDATE()
              WHERE SectionTitle = @sectionTitle`);
    return { action: 'updated', sectionTitle };
  }

  const result = await p.request()
    .input('sectionTitle', sql.NVarChar(500), sectionTitle)
    .input('sectionNumber', sql.Int, sectionNumber)
    .input('htmlContent', sql.NVarChar(sql.MAX), htmlContent)
    .input('canvasContent', sql.NVarChar(sql.MAX), canvasContent)
    .input('apiResponse', sql.NVarChar(sql.MAX), apiResponse)
    .input('navigationConfig', sql.NVarChar(sql.MAX), navigationConfig)
    .query(`INSERT INTO SectionContent (SectionTitle, SectionNumber, HtmlContent, CanvasContent, APIResponse, NavigationConfig)
            VALUES (@sectionTitle, @sectionNumber, @htmlContent, @canvasContent, @apiResponse, @navigationConfig)`);
  return { action: 'inserted', sectionTitle };
}

async function getSectionContent(sectionTitle) {
  const p = await getPool();
  if (!p) return null;

  const result = await p.request()
    .input('sectionTitle', sql.NVarChar(500), sectionTitle)
    .query(`SELECT * FROM SectionContent WHERE SectionTitle = @sectionTitle AND IsActive = 1`);

  return result.recordset[0] || null;
}

async function getAllSectionContents() {
  const p = await getPool();
  if (!p) return [];

  const result = await p.request()
    .query(`SELECT * FROM SectionContent WHERE IsActive = 1 ORDER BY SectionNumber`);

  return result.recordset;
}

async function ensureTable() {
  const p = await getPool();
  if (!p) return;

  const check = await p.request()
    .query(`SELECT OBJECT_ID('SectionContent') AS TableId`);

  if (check.recordset[0].TableId) return;

  await p.request().query(`
    CREATE TABLE SectionContent (
      SectionContentID    INT IDENTITY(1,1) PRIMARY KEY,
      SectionTitle        NVARCHAR(500) NOT NULL,
      SectionNumber       INT NULL,
      HtmlContent         NVARCHAR(MAX) NULL,
      CanvasContent       NVARCHAR(MAX) NULL,
      APIResponse         NVARCHAR(MAX) NULL,
      NavigationConfig    NVARCHAR(MAX) NULL,
      IsActive            BIT DEFAULT 1,
      CreatedAt           DATETIME2 DEFAULT GETUTCDATE(),
      UpdatedAt           DATETIME2 DEFAULT GETUTCDATE(),
      CONSTRAINT UQ_SectionContent_Title UNIQUE (SectionTitle)
    );

    CREATE INDEX IX_SectionContent_IsActive ON SectionContent(IsActive);
  `);

  console.log('  ✓ SectionContent table created');
}

module.exports = {
  getPool,
  upsertSectionContent,
  getSectionContent,
  getAllSectionContents,
  ensureTable,
};
