-- Message Read Receipts
-- Add tracking for individual message read status

CREATE TABLE IF NOT EXISTS community_message_read_receipts (
  id UUID PRIMARY KEY,
  message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
  user_id UUID REFERENCES users(id) ON DELETE CASCADE,
  read_at TIMESTAMP DEFAULT NOW(),
  UNIQUE (message_id, user_id)
);

CREATE INDEX IF NOT EXISTS idx_read_receipts_message ON community_message_read_receipts(message_id);
CREATE INDEX IF NOT EXISTS idx_read_receipts_user ON community_message_read_receipts(user_id);
