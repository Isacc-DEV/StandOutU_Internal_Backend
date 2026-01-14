-- Community Features Schema Additions
-- Run this to add support for file uploads, reactions, replies, pagination, etc.

-- 1. Message attachments and replies
ALTER TABLE community_messages ADD COLUMN IF NOT EXISTS reply_to_message_id UUID REFERENCES community_messages(id);
ALTER TABLE community_messages ADD COLUMN IF NOT EXISTS is_edited BOOLEAN DEFAULT FALSE;
ALTER TABLE community_messages ADD COLUMN IF NOT EXISTS edited_at TIMESTAMP;
ALTER TABLE community_messages ADD COLUMN IF NOT EXISTS is_deleted BOOLEAN DEFAULT FALSE;
ALTER TABLE community_messages ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP;

CREATE INDEX IF NOT EXISTS idx_community_messages_reply ON community_messages(reply_to_message_id);

CREATE TABLE IF NOT EXISTS community_message_attachments (
  id UUID PRIMARY KEY,
  message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
  file_name TEXT NOT NULL,
  file_url TEXT NOT NULL,
  file_size BIGINT NOT NULL,
  mime_type TEXT NOT NULL,
  thumbnail_url TEXT,
  width INTEGER,
  height INTEGER,
  created_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_attachments_message ON community_message_attachments(message_id);

-- 2. Message reactions
CREATE TABLE IF NOT EXISTS community_message_reactions (
  id UUID PRIMARY KEY,
  message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
  user_id UUID REFERENCES users(id) ON DELETE CASCADE,
  emoji TEXT NOT NULL,
  created_at TIMESTAMP DEFAULT NOW(),
  UNIQUE (message_id, user_id, emoji)
);

CREATE INDEX IF NOT EXISTS idx_reactions_message ON community_message_reactions(message_id);
CREATE INDEX IF NOT EXISTS idx_reactions_user ON community_message_reactions(user_id);

-- 3. Unread tracking
CREATE TABLE IF NOT EXISTS community_unread_messages (
  id UUID PRIMARY KEY,
  thread_id UUID REFERENCES community_threads(id) ON DELETE CASCADE,
  user_id UUID REFERENCES users(id) ON DELETE CASCADE,
  last_read_message_id UUID REFERENCES community_messages(id),
  unread_count INTEGER DEFAULT 0,
  last_read_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW(),
  UNIQUE (thread_id, user_id)
);

CREATE INDEX IF NOT EXISTS idx_unread_user ON community_unread_messages(user_id);
CREATE INDEX IF NOT EXISTS idx_unread_thread ON community_unread_messages(thread_id);

-- 4. Channel roles and permissions
ALTER TABLE community_thread_members ADD COLUMN IF NOT EXISTS permissions JSONB DEFAULT '{}'::jsonb;

CREATE TABLE IF NOT EXISTS community_channel_roles (
  id UUID PRIMARY KEY,
  channel_id UUID REFERENCES community_threads(id) ON DELETE CASCADE,
  role_name TEXT NOT NULL,
  permissions JSONB DEFAULT '{}'::jsonb,
  created_at TIMESTAMP DEFAULT NOW(),
  UNIQUE (channel_id, role_name)
);

-- 5. Pinned messages
CREATE TABLE IF NOT EXISTS community_pinned_messages (
  id UUID PRIMARY KEY,
  thread_id UUID REFERENCES community_threads(id) ON DELETE CASCADE,
  message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
  pinned_by UUID REFERENCES users(id),
  pinned_at TIMESTAMP DEFAULT NOW(),
  UNIQUE (thread_id, message_id)
);

CREATE INDEX IF NOT EXISTS idx_pinned_thread ON community_pinned_messages(thread_id);

-- 6. User presence
CREATE TABLE IF NOT EXISTS community_user_presence (
  user_id UUID PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
  status TEXT DEFAULT 'offline',
  last_seen_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);
