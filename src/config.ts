import "dotenv/config";

function normalizeUrl(value: string | undefined) {
  return (value || "").trim().replace(/\/+$/, "");
}

const defaultCorsOrigins = [
  'http://localhost:3000',
  'http://localhost:3300',
  'http://localhost:4000',
  'http://127.0.0.1:3000',
  'http://127.0.0.1:3300',
  'http://127.0.0.1:4000',
  'http://89.117.21.252:3000',
  'http://89.117.21.252:4000',
];

const frontendUrl = normalizeUrl(process.env.FRONTEND_URL || process.env.NEXTAUTH_URL);
const publicApiUrl = normalizeUrl(process.env.PUBLIC_API_URL);
const configuredCorsDefaults = frontendUrl
  ? Array.from(new Set([...defaultCorsOrigins, frontendUrl]))
  : defaultCorsOrigins;

const corsEnv = (process.env.CORS_ORIGINS || '').trim();
const corsOrigins: string[] | true =
  corsEnv === '*'
    ? true
    : corsEnv
        .split(',')
        .map((origin) => origin.trim())
        .filter(Boolean);

export const config = {
  PORT: process.env.PORT ? Number(process.env.PORT) : 4000,
  HOST: (process.env.HOST || '0.0.0.0').trim() || '0.0.0.0',
  
  DATABASE_URL: process.env.DATABASE_URL || 'postgres://postgres:postgres@localhost:5432/ops_db',

  DEBUG_MODE: false,
  
  CORS_ORIGINS: corsEnv ? corsOrigins : configuredCorsDefaults,
  FRONTEND_URL: frontendUrl || 'http://localhost:3000',
  PUBLIC_API_URL: publicApiUrl,
  
  RESUME_DIR: process.env.RESUME_DIR || '',
  
  OPENAI_MODEL: process.env.OPENAI_MODEL || 'gpt-5.4-mini',
  
  SUPABASE_URL: process.env.SUPABASE_URL || '',
  SUPABASE_KEY: process.env.SUPABASE_PUBLISHABLE_KEY || '',
  SUPABASE_BUCKET: process.env.COMMUNITY_FILES_BUCKET_STORAGE || 'community-files',
  
  ENCRYPTION_KEY: process.env.ENCRYPTION_KEY || 'dev-smartwork-change-me',
  
  MS_CLIENT_ID: process.env.MS_CLIENT_ID || '',
  MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET || '',
  MS_TENANT_ID: process.env.MS_TENANT_ID || 'common',
  MS_REDIRECT_URL: process.env.MS_REDIRECT_URL || '',
  
  NEXTAUTH_URL: process.env.NEXTAUTH_URL || '',
  NEXTAUTH_SECRET: process.env.NEXTAUTH_SECRET || '',

  MAIL_SYNC_INTERVAL_MS: process.env.MAIL_SYNC_INTERVAL_MS
    ? Number(process.env.MAIL_SYNC_INTERVAL_MS)
    : 15 * 60 * 1000, // default to 15 minutes
};
