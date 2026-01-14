-- Add JSON storage for parsed resumes.
ALTER TABLE resumes
  ADD COLUMN IF NOT EXISTS resume_json JSONB;
