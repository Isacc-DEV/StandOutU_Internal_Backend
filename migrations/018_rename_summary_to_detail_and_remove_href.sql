-- Rename summary column to detail (if summary exists and detail doesn't)
DO $$
BEGIN
  IF EXISTS (
    SELECT 1 FROM information_schema.columns 
    WHERE table_name = 'tasks' AND column_name = 'summary'
  ) AND NOT EXISTS (
    SELECT 1 FROM information_schema.columns 
    WHERE table_name = 'tasks' AND column_name = 'detail'
  ) THEN
    ALTER TABLE tasks RENAME COLUMN summary TO detail;
  END IF;
END $$;

-- Drop href column if it exists
ALTER TABLE IF EXISTS tasks
  DROP COLUMN IF EXISTS href;
