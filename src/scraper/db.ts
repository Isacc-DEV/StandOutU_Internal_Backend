import { Pool, type PoolConfig } from "pg";
import { env } from "./config";

export type TableNames = {
  countries: string;
  sources: string;
  sourceUrls: string;
  jobLinks: string;
};

type ScrapeDbOptions = {
  pool?: Pool;
  poolConfig?: PoolConfig;
  tableNames?: Partial<TableNames>;
};

export const DEFAULT_TABLE_NAMES: TableNames = {
  countries: "job_countries",
  sources: "job_sources",
  sourceUrls: "job_source_urls",
  jobLinks: "job_links"
};

const ensureSafeTableName = (name: string): string => {
  if (!/^[a-zA-Z0-9_]+$/.test(name)) {
    throw new Error(`Invalid table name: ${name}`);
  }
  return name;
};

export const resolveTableNames = (overrides?: Partial<TableNames>): TableNames => {
  const merged = { ...DEFAULT_TABLE_NAMES, ...overrides };
  return {
    countries: ensureSafeTableName(merged.countries),
    sources: ensureSafeTableName(merged.sources),
    sourceUrls: ensureSafeTableName(merged.sourceUrls),
    jobLinks: ensureSafeTableName(merged.jobLinks)
  };
};

export const buildPoolConfig = (overrides?: PoolConfig): PoolConfig => ({
  host: env.dbHost,
  port: env.dbPort,
  database: env.dbName,
  user: env.dbUser,
  password: env.dbPassword,
  ssl: env.dbSsl ? { rejectUnauthorized: false } : undefined,
  max: env.dbPoolMax,
  idleTimeoutMillis: env.dbPoolIdleMs,
  connectionTimeoutMillis: env.dbPoolConnMs,
  ...overrides
});

export const createScrapePool = (overrides?: PoolConfig): Pool =>
  new Pool(buildPoolConfig(overrides));

export const ensureScrapeSchema = async (
  options: ScrapeDbOptions = {}
): Promise<{ createdTables: string[]; existingTables: string[] }> => {
  const tableNames = resolveTableNames(options.tableNames);
  const pool = options.pool ?? createScrapePool(options.poolConfig);
  const shouldClosePool = !options.pool;
  const expectedTables = [
    tableNames.countries,
    tableNames.sources,
    tableNames.sourceUrls,
    tableNames.jobLinks
  ];

  let transactionStarted = false;
  try {
    console.log('\n📦 [SCRAPER-DB] Initializing scraper database schema...');
    const { rows } = await pool.query<{ table_name: string }>(
      "SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' AND table_name = ANY($1)",
      [expectedTables]
    );

    const existingTables = new Set(rows.map((row) => row.table_name));
    const allPresent = expectedTables.every((name) => existingTables.has(name));
    if (allPresent) {
      console.log('✅ [SCRAPER-DB] All scraper tables already exist');
      return { createdTables: [], existingTables: Array.from(existingTables) };
    }
    const createdTables = expectedTables.filter(
      (name) => !existingTables.has(name)
    );

    console.log(`📝 [SCRAPER-DB] Creating ${createdTables.length} missing table(s)...`);
    await pool.query("BEGIN");
    transactionStarted = true;

    await pool.query(`
      CREATE TABLE IF NOT EXISTS ${tableNames.countries} (
        id SERIAL PRIMARY KEY,
        name VARCHAR(100) NOT NULL
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS ${tableNames.sources} (
        id SERIAL PRIMARY KEY,
        name VARCHAR(100) NOT NULL
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS ${tableNames.sourceUrls} (
        id SERIAL PRIMARY KEY,
        url VARCHAR(200) NOT NULL UNIQUE,
        source_id INTEGER NOT NULL REFERENCES ${tableNames.sources}(id) ON DELETE CASCADE,
        country_id INTEGER NOT NULL REFERENCES ${tableNames.countries}(id) ON DELETE CASCADE
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS ${tableNames.jobLinks} (
        id SERIAL PRIMARY KEY,
        url VARCHAR(200) NOT NULL,
        country_id INTEGER REFERENCES ${tableNames.countries}(id) ON DELETE CASCADE,
        submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `);

    await pool.query(
      `CREATE INDEX IF NOT EXISTS idx_${tableNames.sourceUrls}_source_id ON ${tableNames.sourceUrls} (source_id)`
    );
    await pool.query(
      `CREATE INDEX IF NOT EXISTS idx_${tableNames.sourceUrls}_country_id ON ${tableNames.sourceUrls} (country_id)`
    );
    await pool.query(
      `CREATE INDEX IF NOT EXISTS idx_${tableNames.jobLinks}_country_id ON ${tableNames.jobLinks} (country_id)`
    );

    await pool.query("COMMIT");
    transactionStarted = false;
    console.log('✅ [SCRAPER-DB] Scraper tables created:', createdTables.join(', '));
    console.log('✅ [SCRAPER-DB] Scraper database initialization completed\n');
    return { createdTables, existingTables: Array.from(existingTables) };
  } catch (error) {
    if (transactionStarted) {
      await pool.query("ROLLBACK");
    }
    console.error('❌ [SCRAPER-DB] Scraper database initialization failed:');
    console.error(error);
    throw error;
  } finally {
    if (shouldClosePool) {
      await pool.end();
    }
  }
};
