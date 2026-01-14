-- Add profile_id to resumes for backfilling links.
ALTER TABLE IF EXISTS resumes
  ADD COLUMN IF NOT EXISTS profile_id UUID REFERENCES profiles(id);

-- Optional: set NOT NULL after backfilling if desired.
