import { chromium, type Browser, type Page } from "playwright";
import { XMLParser } from "fast-xml-parser";
import { allowedSourceMap, env } from "./config";
import type { ScrapeStore } from "./types";

type BrowserScrapeConfig = {
  kind: "browser";
  selector: string;
  multiple?: boolean;
  pageSelector?: string;
  pageQuery?: string;
  keywords?: string[];
};

type ApiScrapeConfig = {
  kind: "remotive" | "remoteok" | "weworkremotely";
};

type ScrapeConfig = BrowserScrapeConfig | ApiScrapeConfig;

export type ScrapeSummary = {
  sourcesProcessed: number;
  pagesVisited: number;
  linksFound: number;
  jobLinksCreated: number;
  jobLinksSkipped: number;
  errors: number;
  logEnabled: boolean;
};

const userAgents = [
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15"
];

const sourceConfigs: Record<string, ScrapeConfig> = {
  "Remote Rocketship": {
    kind: "browser",
    selector: "h3.mr-4>a"
  },
  Himalayas: {
    kind: "browser",
    selector: "article.border-gray-200 a.text-xl",
    multiple: true,
    pageSelector: "nav[aria-label=pagination] ul>li:nth-last-child(1)>a",
    pageQuery: "&page=",
    keywords: ["developer", "engineer", "programmer"]
  },
  Remotive: {
    kind: "remotive"
  },
  RemoteOK: {
    kind: "remoteok"
  },
  "We Work Remotely": {
    kind: "weworkremotely"
  }
};

export type ScrapeLogger = (message: string, meta?: Record<string, unknown>) => void;

export type ScrapeOptions = {
  logger?: ScrapeLogger;
  logEnabled?: boolean;
};

const defaultLogger: ScrapeLogger = (message, meta) => {
  const suffix = meta ? ` ${JSON.stringify(meta)}` : "";
  console.log(`[${new Date().toISOString()}] ${message}${suffix}`);
};

const isAllowedSource = (sourceName: string): boolean => {
  const key = sourceName.toLowerCase();
  if (!allowedSourceMap.has(key)) {
    return true;
  }
  return Boolean(allowedSourceMap.get(key));
};

const isTargetClosedError = (error: unknown): boolean => {
  if (!(error instanceof Error)) {
    return false;
  }
  return error.message.toLowerCase().includes("target page, context or browser has been closed");
};

const pickUserAgent = (): string =>
  userAgents[Math.floor(Math.random() * userAgents.length)];

const normalizeProxyServer = (server: string): string =>
  server.includes("://") ? server : `http://${server}`;

const extractNumber = (text: string | null): number => {
  if (!text) {
    return 0;
  }
  const match = text.match(/\d+/);
  return match ? Number.parseInt(match[0], 10) : 0;
};

const matchesKeywords = (url: string, keywords?: string[]): boolean => {
  if (!keywords || keywords.length === 0) {
    return true;
  }
  const normalized = url.toLowerCase();
  return keywords.some((keyword) => normalized.includes(keyword.toLowerCase()));
};

const normalizeJobUrl = (baseUrl: string, href: string): string | null => {
  try {
    const resolved = new URL(href, baseUrl);
    if (!["http:", "https:"].includes(resolved.protocol)) {
      return null;
    }
    return resolved.toString();
  } catch {
    return null;
  }
};

const normalizeApiUrls = (baseUrl: string, urls: string[]): string[] => {
  const uniqueUrls = new Set<string>();
  for (const url of urls) {
    const normalized = normalizeJobUrl(baseUrl, url);
    if (normalized) {
      uniqueUrls.add(normalized);
    }
  }
  return Array.from(uniqueUrls);
};

const fetchJson = async <T>(url: string): Promise<T> => {
  const response = await fetch(url, {
    headers: {
      "User-Agent": pickUserAgent(),
      Accept: "application/json"
    }
  });
  if (!response.ok) {
    throw new Error(`Request failed (${response.status}) for ${url}`);
  }
  return (await response.json()) as T;
};

const fetchText = async (url: string): Promise<string> => {
  const response = await fetch(url, {
    headers: {
      "User-Agent": pickUserAgent(),
      Accept: "application/xml,text/xml"
    }
  });
  if (!response.ok) {
    throw new Error(`Request failed (${response.status}) for ${url}`);
  }
  return response.text();
};

const collectRemotiveJobLinks = async (url: string): Promise<string[]> => {
  const data = await fetchJson<{ jobs?: Array<{ url?: string }> }>(url);
  const urls = (data.jobs ?? [])
    .map((job) => job.url)
    .filter((jobUrl): jobUrl is string => Boolean(jobUrl));
  return normalizeApiUrls(url, urls);
};

const collectRemoteOkJobLinks = async (url: string): Promise<string[]> => {
  const data = await fetchJson<Array<{ url?: string }>>(url);
  const urls = data
    .map((job) => job.url)
    .filter((jobUrl): jobUrl is string => Boolean(jobUrl));
  return normalizeApiUrls(url, urls);
};

const collectWeWorkRemotelyJobLinks = async (url: string): Promise<string[]> => {
  const xml = await fetchText(url);
  const parser = new XMLParser({
    ignoreAttributes: false,
    trimValues: true
  });
  const parsed = parser.parse(xml) as {
    rss?: { channel?: { item?: Array<{ link?: string }> | { link?: string } } };
  };
  const items = parsed.rss?.channel?.item ?? [];
  const list = Array.isArray(items) ? items : [items];
  const urls = list
    .map((item) => item.link)
    .filter((jobUrl): jobUrl is string => Boolean(jobUrl));
  return normalizeApiUrls(url, urls);
};

const collectApiJobLinks = async (
  url: string,
  config: ApiScrapeConfig
): Promise<string[]> => {
  switch (config.kind) {
    case "remotive":
      return collectRemotiveJobLinks(url);
    case "remoteok":
      return collectRemoteOkJobLinks(url);
    case "weworkremotely":
      return collectWeWorkRemotelyJobLinks(url);
    default:
      return [];
  }
};

const createBrowser = async (): Promise<Browser> => {
  const proxyServer = env.proxyServer.trim();
  if (!proxyServer) {
    return chromium.launch({ headless: env.scrapeHeadless });
  }

  return chromium.launch({
    headless: env.scrapeHeadless,
    proxy: {
      server: normalizeProxyServer(proxyServer),
      username: env.proxyUsername || undefined,
      password: env.proxyPassword || undefined
    }
  });
};

const collectJobLinks = async (
  page: Page,
  pageUrl: string,
  config: BrowserScrapeConfig,
  options?: { skipNavigation?: boolean }
): Promise<string[]> => {
  if (!options?.skipNavigation) {
    await page.goto(pageUrl, {
      waitUntil: "domcontentloaded",
      timeout: env.scrapeTimeoutMs
    });
  }

  await page.waitForSelector(config.selector, {
    timeout: env.scrapeTimeoutMs
  });

  const hrefs = await page.$$eval(config.selector, (nodes) =>
    nodes
      .map((node) => node.getAttribute("href"))
      .filter((href): href is string => Boolean(href))
  );

  const uniqueUrls = new Set<string>();
  for (const href of hrefs) {
    const jobUrl = normalizeJobUrl(pageUrl, href);
    if (!jobUrl) {
      continue;
    }
    if (!matchesKeywords(jobUrl, config.keywords)) {
      continue;
    }
    uniqueUrls.add(jobUrl);
  }

  return Array.from(uniqueUrls);
};

const resolvePageUrls = async (
  page: Page,
  url: string,
  config: BrowserScrapeConfig
): Promise<{ pageUrls: string[]; firstPageReady: boolean }> => {
  if (!config.multiple || !config.pageSelector) {
    return { pageUrls: [url], firstPageReady: false };
  }

  let pageCount = 1;
  try {
    await page.goto(url, {
      waitUntil: "domcontentloaded",
      timeout: env.scrapeTimeoutMs
    });
    await page.waitForSelector(config.selector, {
      timeout: env.scrapeTimeoutMs
    });
    const pageText = await page.textContent(config.pageSelector);
    const extracted = extractNumber(pageText);
    const pageLimit = Math.max(1, env.scrapeMaxPages);
    if (extracted > 1) {
      pageCount = Math.min(extracted, pageLimit);
    }
  } catch {
    return { pageUrls: [url], firstPageReady: false };
  }

  const pageQuery = config.pageQuery ?? "&page=";
  const pageUrls = Array.from({ length: pageCount }, (_value, index) =>
    index === 0 ? url : `${url}${pageQuery}${index + 1}`
  );

  return { pageUrls, firstPageReady: true };
};

export const scrapeUrls = async (
  scrapeStore: ScrapeStore,
  options: ScrapeOptions = {}
): Promise<ScrapeSummary> => {
  const logEnabled = options.logEnabled ?? env.scrapeLogEnabled;
  const logger = options.logger ?? defaultLogger;
  const log: ScrapeLogger = logEnabled ? logger : () => {};
  const summary: ScrapeSummary = {
    sourcesProcessed: 0,
    pagesVisited: 0,
    linksFound: 0,
    jobLinksCreated: 0,
    jobLinksSkipped: 0,
    errors: 0,
    logEnabled
  };

  log("Scrape started");
  let browser: Browser | null = null;
  const getBrowser = async (): Promise<Browser> => {
    if (!browser) {
      browser = await createBrowser();
      log("Browser launched");
    }
    return browser;
  };
  try {
    const sourceUrls = await scrapeStore.listSourceUrls();
    log("Loaded source urls", { count: sourceUrls.length });
    for (const sourceUrl of sourceUrls) {
      const config = sourceConfigs[sourceUrl.sourceName];
      if (!config) {
        log("Skipping source (no config)", {
          source: sourceUrl.sourceName,
          url: sourceUrl.url
        });
        continue;
      }
      if (!isAllowedSource(sourceUrl.sourceName)) {
        log("Skipping source (not allowed)", {
          source: sourceUrl.sourceName,
          url: sourceUrl.url
        });
        continue;
      }

      summary.sourcesProcessed += 1;
      const sourceStats = {
        created: 0,
        skipped: 0,
        pages: 0,
        links: 0
      };
      log("Scraping source", {
        source: sourceUrl.sourceName,
        url: sourceUrl.url
      });

      if (config.kind !== "browser") {
        try {
          const jobUrls = await collectApiJobLinks(sourceUrl.url, config);
          summary.pagesVisited += 1;
          sourceStats.pages += 1;
          summary.linksFound += jobUrls.length;
          sourceStats.links += jobUrls.length;

          for (const jobUrl of jobUrls) {
            const created = await scrapeStore.createJobLink(
              jobUrl,
              sourceUrl.countryId
            );
            if (created) {
              summary.jobLinksCreated += 1;
              sourceStats.created += 1;
            } else {
              summary.jobLinksSkipped += 1;
              sourceStats.skipped += 1;
            }
          }

          log("Scrape success", {
            source: sourceUrl.sourceName,
            url: sourceUrl.url,
            pages: sourceStats.pages,
            links: sourceStats.links,
            created: sourceStats.created,
            skipped: sourceStats.skipped
          });
        } catch (error) {
          summary.errors += 1;
          log("Scrape failed", {
            source: sourceUrl.sourceName,
            url: sourceUrl.url,
            error: error instanceof Error ? error.message : String(error)
          });
        }
        continue;
      }

      const activeBrowser = await getBrowser();
      let context = await activeBrowser.newContext({
        userAgent: pickUserAgent()
      });
      let page = await context.newPage();

      const recreateContext = async (): Promise<void> => {
        try {
          await page.close();
        } catch {
          // Ignore close failures for already-closed pages.
        }
        try {
          await context.close();
        } catch {
          // Ignore close failures for already-closed contexts.
        }
        if (!browser) {
          throw new Error("browser is not defined")
        }

        context = await browser.newContext({
          userAgent: pickUserAgent()
        });
        page = await context.newPage();
      };

      try {
        const { pageUrls, firstPageReady } = await resolvePageUrls(
          page,
          sourceUrl.url,
          config
        );
        summary.pagesVisited += pageUrls.length;
        sourceStats.pages += pageUrls.length;

        for (const [index, pageUrl] of pageUrls.entries()) {
          const skipNavigation = firstPageReady && index === 0;
          let jobUrls: string[] = [];

          try {
            jobUrls = await collectJobLinks(page, pageUrl, config, {
              skipNavigation
            });
          } catch (error) {
            summary.errors += 1;
            log("Page scrape failed", {
              source: sourceUrl.sourceName,
              url: pageUrl,
              error: error instanceof Error ? error.message : String(error)
            });

            if (isTargetClosedError(error)) {
              log("Page closed, recreating context", { url: pageUrl });
              await recreateContext();
              try {
                jobUrls = await collectJobLinks(page, pageUrl, config);
              } catch (retryError) {
                summary.errors += 1;
                log("Page scrape retry failed", {
                  source: sourceUrl.sourceName,
                  url: pageUrl,
                  error:
                    retryError instanceof Error
                      ? retryError.message
                      : String(retryError)
                });
                continue;
              }
            } else {
              continue;
            }
          }

          summary.linksFound += jobUrls.length;
          sourceStats.links += jobUrls.length;

          for (const jobUrl of jobUrls) {
            const created = await scrapeStore.createJobLink(
              jobUrl,
              sourceUrl.countryId
            );
            if (created) {
              summary.jobLinksCreated += 1;
              sourceStats.created += 1;
            } else {
              summary.jobLinksSkipped += 1;
              sourceStats.skipped += 1;
            }
          }
        }
        log("Scrape success", {
          source: sourceUrl.sourceName,
          url: sourceUrl.url,
          pages: sourceStats.pages,
          links: sourceStats.links,
          created: sourceStats.created,
          skipped: sourceStats.skipped
        });
      } catch (error) {
        summary.errors += 1;
        log("Scrape failed", {
          source: sourceUrl.sourceName,
          url: sourceUrl.url,
          error: error instanceof Error ? error.message : String(error)
        });
      } finally {
        try {
          await page.close();
        } catch {
          // Ignore close failures for already-closed pages.
        }
        try {
          await context.close();
        } catch {
          // Ignore close failures for already-closed contexts.
        }
      }
    }
  } finally {
    if (browser) {
      await browser.close();
    }
  }

  return summary;
};
