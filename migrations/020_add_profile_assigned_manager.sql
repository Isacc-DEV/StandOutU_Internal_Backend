ALTER TABLE IF EXISTS profiles
  ADD COLUMN IF NOT EXISTS assigned_manager_user_id UUID REFERENCES users(id);

UPDATE profiles
SET assigned_manager_user_id = created_by
WHERE assigned_manager_user_id IS NULL
  AND created_by IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_profiles_assigned_manager_user_id
  ON profiles(assigned_manager_user_id);
