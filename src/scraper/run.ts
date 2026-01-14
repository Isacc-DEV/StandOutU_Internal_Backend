import cron from "node-cron";
import { scrapeUrls, type ScrapeLogger } from "./scrape";
import { ensureScrapeSchema, createScrapePool } from "./db";
import { createPostgresScrapeStore } from "./postgres";
import { env } from "./config";

export type ScraperSchedulerOptions = {
  logEnabled?: boolean;
  logger?: ScrapeLogger;
  cronExpression?: string;
  runOnStart?: boolean;
};

const defaultLogger: ScrapeLogger = (message, meta) => {
  const suffix = meta ? ` ${JSON.stringify(meta)}` : "";
  console.log(`[${new Date().toISOString()}] ${message}${suffix}`);
};

const oneHourMs = 60 * 60 * 1000;
const defaultCronExpression = "*/1 * * * *";

let isRunning = false;
let lastCompletedAt: Date | null = null;

const runOnce = async (options: Required<ScraperSchedulerOptions>): Promise<void> => {
  const log = options.logEnabled ? options.logger : () => {};
  const pool = createScrapePool();
  const store = createPostgresScrapeStore({ pool });
  try {
    const schemaResult = await ensureScrapeSchema({ pool });
    if (schemaResult.createdTables.length > 0) {
      log("Database schema created", {
        tables: schemaResult.createdTables
      });
    } else {
      log("Database schema already exists");
    }
    const summary = await scrapeUrls(store, {
      logEnabled: options.logEnabled,
      logger: options.logger
    });
    log("Scrape complete", summary);
  } finally {
    await store.close();
    await pool.end();
  }
};

const runWithCooldown = async (
  options: Required<ScraperSchedulerOptions>
): Promise<void> => {
  const log = options.logEnabled ? options.logger : () => {};
  if (isRunning) {
    return;
  }

  const now = Date.now();
  if (lastCompletedAt && now - lastCompletedAt.getTime() < oneHourMs) {
    return;
  }

  isRunning = true;
  try {
    await runOnce(options);
  } catch (error) {
    log("Scrape failed", {
      error: error instanceof Error ? error.message : String(error)
    });
  } finally {
    lastCompletedAt = new Date();
    isRunning = false;
    log("Next scrape eligible", {
      at: new Date(lastCompletedAt.getTime() + oneHourMs).toISOString()
    });
  }
};

export const startScraperScheduler = async (
  options: ScraperSchedulerOptions = {}
): Promise<void> => {
  const resolvedOptions: Required<ScraperSchedulerOptions> = {
    logEnabled: options.logEnabled ?? env.scrapeLogEnabled,
    logger: options.logger ?? defaultLogger,
    cronExpression: options.cronExpression ?? defaultCronExpression,
    runOnStart: options.runOnStart ?? true
  };
  const log = resolvedOptions.logEnabled ? resolvedOptions.logger : () => {};

  if (resolvedOptions.runOnStart) {
    await runWithCooldown(resolvedOptions);
  }
  cron.schedule(resolvedOptions.cronExpression, () => {
    void runWithCooldown(resolvedOptions);
  });
  log("Cron schedule started", {
    cronExpression: resolvedOptions.cronExpression
  });
};

if (require.main === module) {
  startScraperScheduler().catch((error) => {
    console.error("Scrape scheduler failed:", error);
    process.exit(1);
  });
}
