import "dotenv/config";

const defaultCorsOrigins = [
  'http://localhost:3000',
  'http://localhost:3300',
  'http://localhost:4000',
  'http://127.0.0.1:3000',
  'http://127.0.0.1:3300',
  'http://127.0.0.1:4000',
  'http://38.79.97.159:3000'
];

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
  
  DATABASE_URL: process.env.DATABASE_URL || 'postgres://postgres:postgres@localhost:5432/ops_db',

  DEBUG_MODE: false,
  
  CORS_ORIGINS: corsEnv ? corsOrigins : defaultCorsOrigins,
  
  RESUME_DIR: process.env.RESUME_DIR || '',
  
  HF_TOKEN: process.env.HF_TOKEN || 
    process.env.HUGGINGFACEHUB_API_TOKEN || 
    process.env.HUGGING_FACE_TOKEN || 
    '',
  
  HF_MODEL: process.env.HF_MODEL || 'meta-llama/Meta-Llama-3-8B-Instruct',
  
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
};
