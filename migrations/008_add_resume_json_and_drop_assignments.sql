-- Ensure resume_json column exists.
ALTER TABLE IF EXISTS resumes
  ADD COLUMN IF NOT EXISTS resume_json JSONB;

-- Remove assignments table (no longer used).
DROP TABLE IF EXISTS assignments;
