-- Add mailbox_id column to calendar_events table
ALTER TABLE IF EXISTS calendar_events
  ADD COLUMN IF NOT EXISTS mailbox_id UUID;

-- Make mailbox column nullable (since we're using mailbox_id now)
DO $$
BEGIN
  -- Check if mailbox column has NOT NULL constraint
  IF EXISTS (
    SELECT 1 FROM information_schema.columns 
    WHERE table_name = 'calendar_events' 
    AND column_name = 'mailbox' 
    AND is_nullable = 'NO'
  ) THEN
    ALTER TABLE calendar_events ALTER COLUMN mailbox DROP NOT NULL;
  END IF;
END $$;

-- Add foreign key constraint (drop first if exists to avoid errors)
DO $$
BEGIN
  IF EXISTS (
    SELECT 1 FROM information_schema.table_constraints 
    WHERE constraint_name = 'fk_calendar_events_mailbox_id' 
    AND table_name = 'calendar_events'
  ) THEN
    ALTER TABLE calendar_events DROP CONSTRAINT fk_calendar_events_mailbox_id;
  END IF;
END $$;

ALTER TABLE IF EXISTS calendar_events
  ADD CONSTRAINT fk_calendar_events_mailbox_id
  FOREIGN KEY (mailbox_id) REFERENCES user_oauth_accounts(id) ON DELETE SET NULL;

-- Migrate existing data: Match mailbox email to user_oauth_accounts.email and set mailbox_id
UPDATE calendar_events ce
SET mailbox_id = (
  SELECT uoa.id
  FROM user_oauth_accounts uoa
  WHERE LOWER(uoa.email) = LOWER(ce.mailbox)
  LIMIT 1
)
WHERE mailbox_id IS NULL;

-- Add index for mailbox_id
CREATE INDEX IF NOT EXISTS idx_calendar_events_mailbox_id ON calendar_events(mailbox_id);

-- Update existing index to include mailbox_id for better query performance
DROP INDEX IF EXISTS idx_calendar_events_owner_mailbox;
CREATE INDEX IF NOT EXISTS idx_calendar_events_owner_mailbox ON calendar_events(owner_user_id, mailbox_id);

-- Add index for owner_user_id and mailbox_id together for filtering
CREATE INDEX IF NOT EXISTS idx_calendar_events_owner_mailbox_id ON calendar_events(owner_user_id, mailbox_id);
