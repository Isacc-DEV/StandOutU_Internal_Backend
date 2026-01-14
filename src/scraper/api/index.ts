import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { forbidObserver } from "../../auth";
import { createScrapePool, ensureScrapeSchema, resolveTableNames } from "../db";
import { createScraperApiStore } from "./store";

const querySchema = z.object({
  limit: z.coerce.number().int().min(1).max(200).optional(),
  offset: z.coerce.number().int().min(0).optional(),
  countryId: z.coerce.number().int().min(1).optional(),
  search: z.string().trim().max(300).optional(),
  since: z.string().trim().optional(),
  until: z.string().trim().optional()
});

const normalizeOptionalString = (value?: string) => {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
};

const normalizeDateInput = (value?: string) => {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!trimmed) return undefined;
  const timestamp = Date.parse(trimmed);
  if (Number.isNaN(timestamp)) return null;
  return new Date(timestamp).toISOString();
};

export const registerScraperApiRoutes = async (
  app: FastifyInstance
): Promise<void> => {
  const pool = createScrapePool();
  const tableNames = resolveTableNames();
  await ensureScrapeSchema({ pool, tableNames });
  const store = createScraperApiStore(pool, tableNames);

  app.get("/scraper/job-links", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    const parsed = querySchema.safeParse(request.query ?? {});
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }

    const limit = parsed.data.limit ?? 50;
    const offset = parsed.data.offset ?? 0;
    const countryId = parsed.data.countryId;
    const search = normalizeOptionalString(parsed.data.search);
    const since = normalizeDateInput(parsed.data.since);
    if (since === null) {
      return reply.status(400).send({ message: "Invalid since date" });
    }
    const until = normalizeDateInput(parsed.data.until);
    if (until === null) {
      return reply.status(400).send({ message: "Invalid until date" });
    }

    const result = await store.listJobLinks({
      limit,
      offset,
      countryId,
      search,
      since: since ?? undefined,
      until: until ?? undefined
    });

    return {
      ...result,
      limit,
      offset
    };
  });

  app.get("/scraper/countries", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    const countries = await store.listCountries();
    return { countries };
  });
};
