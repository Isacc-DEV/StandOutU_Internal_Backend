CREATE TABLE IF NOT EXISTS profile_accounts (
  id UUID PRIMARY KEY,
  profile_id UUID REFERENCES profiles(id) ON DELETE CASCADE,
  provider TEXT NOT NULL DEFAULT 'MICROSOFT',
  email TEXT NOT NULL,
  display_name TEXT,
  timezone TEXT DEFAULT 'UTC',
  status TEXT DEFAULT 'ACTIVE',
  last_sync_at TIMESTAMP,
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),
  UNIQUE (profile_id, email)
);

CREATE INDEX IF NOT EXISTS idx_profile_accounts_profile ON profile_accounts(profile_id);
