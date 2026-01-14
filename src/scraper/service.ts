import { env } from "./config";
import { startScraperScheduler, type ScraperSchedulerOptions } from "./run";
import { seedScrapeData } from "./seed";
import type { ScrapeLogger } from "./scrape";

type ScraperServiceOptions = {
  logEnabled?: boolean;
  logger?: ScrapeLogger;
  runOnStart?: boolean;
  cronExpression?: string;
};

const defaultLogger: ScrapeLogger = (message, meta) => {
  const suffix = meta ? ` ${JSON.stringify(meta)}` : "";
  console.log(`[${new Date().toISOString()}] ${message}${suffix}`);
};

export const startScraperService = async (
  options: ScraperServiceOptions = {}
): Promise<void> => {
  const logEnabled = options.logEnabled ?? env.scrapeLogEnabled;
  const logger = options.logger ?? defaultLogger;
  const log: ScrapeLogger = logEnabled ? logger : () => {};

  try {
    log("Scrape seed starting");
    await seedScrapeData({ logEnabled, logger });
  } catch (error) {
    log("Scrape seed failed", {
      error: error instanceof Error ? error.message : String(error)
    });
  }

  const schedulerOptions: ScraperSchedulerOptions = {
    logEnabled,
    logger,
    cronExpression: options.cronExpression,
    runOnStart: options.runOnStart
  };

  try {
    await startScraperScheduler(schedulerOptions);
  } catch (error) {
    log("Scrape scheduler failed", {
      error: error instanceof Error ? error.message : String(error)
    });
  }
};
