import type { Pool } from "pg";
import { resolveTableNames, type TableNames } from "../db";

export type JobLinkRecord = {
  id: number;
  url: string;
  countryId: number | null;
  countryName: string | null;
  submittedAt: string;
};

export type CountryRecord = {
  id: number;
  name: string;
};

export type ListJobLinksInput = {
  limit: number;
  offset: number;
  countryId?: number | null;
  search?: string;
  since?: string;
  until?: string;
};

export type ListJobLinksResult = {
  total: number;
  items: JobLinkRecord[];
};

const buildJobLinksWhere = (filters: ListJobLinksInput) => {
  const clauses: string[] = [];
  const values: Array<string | number> = [];

  if (filters.countryId === null) {
    clauses.push("jl.country_id IS NULL");
  } else if (typeof filters.countryId === "number") {
    values.push(filters.countryId);
    clauses.push(`jl.country_id = $${values.length}`);
  }

  if (filters.search) {
    values.push(`%${filters.search}%`);
    clauses.push(`jl.url ILIKE $${values.length}`);
  }

  if (filters.since) {
    values.push(filters.since);
    clauses.push(`jl.submitted_at >= $${values.length}`);
  }

  if (filters.until) {
    values.push(filters.until);
    clauses.push(`jl.submitted_at <= $${values.length}`);
  }

  return {
    whereSql: clauses.length ? `WHERE ${clauses.join(" AND ")}` : "",
    values
  };
};

export const createScraperApiStore = (
  pool: Pool,
  tableNames: TableNames = resolveTableNames()
) => {
  const listJobLinks = async (
    filters: ListJobLinksInput
  ): Promise<ListJobLinksResult> => {
    const { whereSql, values } = buildJobLinksWhere(filters);
    const countQuery = `
      SELECT COUNT(*)::int AS total
      FROM ${tableNames.jobLinks} jl
      ${whereSql}
    `;
    const countResult = await pool.query<{ total: number }>(countQuery, values);
    const total = countResult.rows[0]?.total ?? 0;

    const listValues = [...values, filters.limit, filters.offset];
    const listQuery = `
      SELECT
        jl.id,
        jl.url,
        jl.country_id AS "countryId",
        c.name AS "countryName",
        jl.submitted_at AS "submittedAt"
      FROM ${tableNames.jobLinks} jl
      LEFT JOIN ${tableNames.countries} c ON jl.country_id = c.id
      ${whereSql}
      ORDER BY jl.submitted_at DESC, jl.id DESC
      LIMIT $${listValues.length - 1}
      OFFSET $${listValues.length}
    `;
    const { rows } = await pool.query<JobLinkRecord>(listQuery, listValues);
    return { total, items: rows };
  };

  const listCountries = async (): Promise<CountryRecord[]> => {
    const query = `
      SELECT id, name
      FROM ${tableNames.countries}
      ORDER BY name ASC
    `;
    const { rows } = await pool.query<CountryRecord>(query);
    return rows;
  };

  return { listJobLinks, listCountries };
};
