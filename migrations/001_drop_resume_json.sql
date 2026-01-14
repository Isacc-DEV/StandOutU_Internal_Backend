-- Drop the old JSON column now that resumes are stored as files.
ALTER TABLE IF EXISTS resumes
  DROP COLUMN IF EXISTS resume_json;
