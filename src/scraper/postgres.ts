import { Pool, type PoolConfig } from "pg";
import type { ScrapeStore, SourceUrlRecord } from "./types";
import { buildPoolConfig } from "./db";

type TableNames = {
  sourceUrls: string;
  sources: string;
  jobLinks: string;
};

type PostgresStoreOptions = {
  pool?: Pool;
  tableNames?: Partial<TableNames>;
  poolConfig?: PoolConfig;
};

const defaultTableNames: TableNames = {
  sourceUrls: "job_source_urls",
  sources: "job_sources",
  jobLinks: "job_links"
};

const ensureSafeTableName = (name: string): string => {
  if (!/^[a-zA-Z0-9_]+$/.test(name)) {
    throw new Error(`Invalid table name: ${name}`);
  }
  return name;
};

const normalizeTableNames = (overrides?: Partial<TableNames>): TableNames => {
  const merged = { ...defaultTableNames, ...overrides };
  return {
    sourceUrls: ensureSafeTableName(merged.sourceUrls),
    sources: ensureSafeTableName(merged.sources),
    jobLinks: ensureSafeTableName(merged.jobLinks)
  };
};

export const createPostgresScrapeStore = (
  options: PostgresStoreOptions = {}
): ScrapeStore => {
  const tableNames = normalizeTableNames(options.tableNames);
  const pool = options.pool ?? new Pool(buildPoolConfig(options.poolConfig));
  const shouldClosePool = !options.pool;

  const listSourceUrls = async (): Promise<SourceUrlRecord[]> => {
    const query = `
      SELECT su.id,
             su.url,
             su.country_id AS "countryId",
             s.name AS "sourceName"
      FROM ${tableNames.sourceUrls} su
      JOIN ${tableNames.sources} s ON su.source_id = s.id
      ORDER BY su.id
    `;
    const { rows } = await pool.query<SourceUrlRecord>(query);
    return rows;
  };

  const createJobLink = async (
    url: string,
    countryId: number | null
  ): Promise<boolean> => {
    const query = `
      INSERT INTO ${tableNames.jobLinks} (url, country_id, submitted_at)
      SELECT $1::varchar, $2::int, NOW()
      WHERE NOT EXISTS (
        SELECT 1 FROM ${tableNames.jobLinks} WHERE url = $1::varchar
      )
    `;
    const result = await pool.query(query, [url, countryId]);
    return (result.rowCount ?? 0) > 0;
  };

  const close = async (): Promise<void> => {
    if (shouldClosePool) {
      await pool.end();
    }
  };

  return { listSourceUrls, createJobLink, close };
};
