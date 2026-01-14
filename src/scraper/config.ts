import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";

const envPath =
  process.env.SCRAPER_ENV_PATH ?? path.resolve(process.cwd(), ".env");

if (fs.existsSync(envPath)) {
  dotenv.config({ path: envPath });
} else {
  dotenv.config();
}

const toNumber = (value: string | undefined, fallback: number): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};

const toBoolean = (value: string | undefined, fallback: boolean): boolean => {
  if (value === undefined) {
    return fallback;
  }
  return ["true", "1", "yes"].includes(value.toLowerCase());
};

export const env = {
  scrapeHeadless: toBoolean(process.env.SCRAPE_HEADLESS, true),
  scrapeTimeoutMs: toNumber(process.env.SCRAPE_TIMEOUT_MS, 20000),
  scrapeMaxPages: toNumber(process.env.SCRAPE_MAX_PAGES, 2),
  proxyServer: process.env.PROXY_SERVER ?? "",
  proxyUsername: process.env.PROXY_USERNAME ?? "",
  proxyPassword: process.env.PROXY_PASSWORD ?? "",
  dbHost: process.env.DB_HOST ?? "localhost",
  dbPort: toNumber(process.env.DB_PORT, 5432),
  dbName: process.env.DB_NAME ?? "ops_db",
  dbUser: process.env.DB_USER ?? "postgres",
  dbPassword: process.env.DB_PASSWORD ?? "postgres",
  dbSsl: toBoolean(process.env.DB_SSL, false),
  dbPoolMax: toNumber(process.env.DB_POOL_MAX, 10),
  dbPoolIdleMs: toNumber(process.env.DB_POOL_IDLE_MS, 30000),
  dbPoolConnMs: toNumber(process.env.DB_POOL_CONN_MS, 2000),
  scrapeLogEnabled: toBoolean(process.env.SCRAPE_LOG_ENABLED, false)
};

export const allowedSources: Record<string, boolean> = {
  "Remote Rocketship": true,
  Himalayas: true,
  Remotive: true,
  RemoteOK: true,
  "We Work Remotely": true
};

export const allowedSourceMap = new Map(
  Object.entries(allowedSources).map(([key, value]) => [
    key.toLowerCase(),
    value
  ])
);
