import { randomUUID } from 'crypto';
import { Pool, PoolClient } from 'pg';
import {
  ApplicationRecord,
  ApplicationSummary,
  Assignment,
  DailyReport,
  DailyReportAttachment,
  DailyReportWithUser,
  Notification,
  CommunityMessage,
  CommunityMessageExtended,
  CommunityThread,
  CommunityThreadSummary,
  LabelAlias,
  MessageAttachment,
  MessageReaction,
  PinnedMessage,
  Profile,
  ProfileAccount,
  ProfileAccountWithProfile,
  Task,
  TaskAssignmentRequestDetail,
  TaskAssignmentRequestStatus,
  TaskAssignee,
  TaskDoneRequestDetail,
  TaskDoneRequestStatus,
  ReactionSummary,
  ResumeTemplate,
  UnreadInfo,
  User,
  UserOAuthAccount,
  UserPresence,
} from './types';
import {
  APPLICATION_SUCCESS_DEFAULTS,
  APPLICATION_SUCCESS_KEY,
  normalizeLabelAlias,
} from './labelAliases';
import { config } from './config';

type CalendarEventInput = {
  id: string;
  mailboxId: string;  // Changed from mailbox: string
  mailbox?: string;  // Keep for backward compatibility during transition
  title?: string;
  start: string;
  end: string;
  isAllDay?: boolean;
  organizer?: string;
  location?: string;
};

type StoredCalendarEvent = {
  id: string;
  mailboxId: string | null;  // Changed from mailbox: string
  mailbox?: string;  // Keep for backward compatibility during transition
  title: string;
  start: string;
  end: string;
  isAllDay?: boolean;
  organizer?: string;
  location?: string;
  timezone?: string | null;
};

const DEFAULT_ADMIN_DATABASE = 'postgres';

function getDatabaseName(connectionString: string): string | null {
  try {
    const url = new URL(connectionString);
    const pathname = url.pathname.replace(/^\/+/, '');
    if (!pathname) return null;
    return decodeURIComponent(pathname);
  } catch {
    return null;
  }
}

function buildAdminConnectionString(
  connectionString: string,
  adminDatabase: string,
): string | null {
  try {
    const url = new URL(connectionString);
    url.pathname = `/${adminDatabase}`;
    return url.toString();
  } catch {
    return null;
  }
}

function quoteIdentifier(identifier: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(identifier)) return identifier;
  return `"${identifier.replace(/"/g, '""')}"`;
}

async function ensureDatabaseExists(connectionString: string) {
  const databaseName = getDatabaseName(connectionString);
  if (!databaseName) return;
  if (databaseName === DEFAULT_ADMIN_DATABASE || databaseName === 'template1') return;
  const adminConnectionString = buildAdminConnectionString(
    connectionString,
    DEFAULT_ADMIN_DATABASE,
  );
  if (!adminConnectionString) return;
  const adminPool = new Pool({ connectionString: adminConnectionString });
  try {
    const { rows } = await adminPool.query<{ exists: number }>(
      'SELECT 1 AS exists FROM pg_database WHERE datname = $1',
      [databaseName],
    );
    if (rows.length > 0) return;
    const dbName = quoteIdentifier(databaseName);
    try {
      await adminPool.query(`CREATE DATABASE ${dbName}`);
    } catch (err) {
      const code = (err as { code?: string } | undefined)?.code;
      if (code !== '42P04') {
        throw err;
      }
    }
  } finally {
    await adminPool.end();
  }
}

export const pool = new Pool({
  connectionString: config.DATABASE_URL,
});

export async function initDb() {
  console.log('\n📦 [DB] Starting database initialization...\n');
  let client: PoolClient;
  try {
    console.log('🔌 [DB] Connecting to database...');
    client = await pool.connect();
    console.log('✅ [DB] Database connection established');
  } catch (err) {
    const code = (err as { code?: string } | undefined)?.code;
    if (code !== '3D000') {
      console.error('❌ [DB] Database connection failed:', err);
      throw err;
    }
    console.log('📝 [DB] Database does not exist, creating...');
    await ensureDatabaseExists(config.DATABASE_URL);
    console.log('✅ [DB] Database created successfully');
    client = await pool.connect();
  }
  try {
    console.log('\n📋 [DB] Creating tables...');
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id UUID PRIMARY KEY,
        email TEXT UNIQUE NOT NULL,
        user_name TEXT UNIQUE NOT NULL,
        name TEXT NOT NULL,
        avatar_url TEXT,
        bio TEXT,
        role TEXT NOT NULL,
        password TEXT NOT NULL,
        is_active BOOLEAN DEFAULT TRUE,
        created_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS profiles (
        id UUID PRIMARY KEY,
        display_name TEXT NOT NULL,
        base_info JSONB DEFAULT '{}'::jsonb,
        base_resume JSONB DEFAULT '{}'::jsonb,
        base_additional_bullets JSONB DEFAULT '{}'::jsonb,
        created_by UUID,
        assigned_bidder_id UUID REFERENCES users(id),
        assigned_by UUID REFERENCES users(id),
        assigned_at TIMESTAMP,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS resume_templates (
        id UUID PRIMARY KEY,
        name TEXT NOT NULL,
        description TEXT,
        html TEXT NOT NULL,
        created_by UUID REFERENCES users(id),
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS applications (
        id UUID PRIMARY KEY,
        session_id UUID UNIQUE,
        bidder_user_id UUID,
        profile_id UUID,
        resume_id UUID,
        url TEXT,
        domain TEXT,
        company TEXT,
        status TEXT DEFAULT 'in_review',
        is_reviewed BOOLEAN DEFAULT FALSE,
        reviewed_by UUID REFERENCES users(id),
        reviewed_at TIMESTAMP,
        created_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS tasks (
        id UUID PRIMARY KEY,
        title TEXT NOT NULL,
        detail TEXT,
        status TEXT NOT NULL,
        priority TEXT NOT NULL,
        due_date DATE,
        project TEXT,
        notes TEXT,
        tags TEXT[] DEFAULT ARRAY[]::text[],
        created_by UUID REFERENCES users(id),
        rejected_by UUID REFERENCES users(id),
        rejected_at TIMESTAMP,
        rejection_reason TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS task_assignees (
        id UUID PRIMARY KEY,
        task_id UUID REFERENCES tasks(id) ON DELETE CASCADE,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        assigned_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (task_id, user_id)
      );

      CREATE TABLE IF NOT EXISTS task_assignment_requests (
        id UUID PRIMARY KEY,
        task_id UUID REFERENCES tasks(id) ON DELETE CASCADE,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        requested_by UUID REFERENCES users(id),
        status TEXT NOT NULL,
        reviewed_by UUID REFERENCES users(id),
        reviewed_at TIMESTAMP,
        rejection_reason TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (task_id, user_id)
      );

      CREATE TABLE IF NOT EXISTS task_done_requests (
        id UUID PRIMARY KEY,
        task_id UUID REFERENCES tasks(id) ON DELETE CASCADE,
        requested_by UUID REFERENCES users(id),
        status TEXT NOT NULL,
        reviewed_by UUID REFERENCES users(id),
        reviewed_at TIMESTAMP,
        rejection_reason TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS label_aliases (
        id UUID PRIMARY KEY,
        canonical_key TEXT NOT NULL,
        alias TEXT NOT NULL,
        normalized_alias TEXT NOT NULL UNIQUE,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

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

      CREATE TABLE IF NOT EXISTS community_threads (
        id UUID PRIMARY KEY,
        thread_type TEXT NOT NULL,
        name TEXT,
        name_key TEXT UNIQUE,
        description TEXT,
        created_by UUID REFERENCES users(id),
        is_private BOOLEAN DEFAULT FALSE,
        created_at TIMESTAMP DEFAULT NOW(),
        last_message_at TIMESTAMP
      );

      CREATE TABLE IF NOT EXISTS community_thread_members (
        id UUID PRIMARY KEY,
        thread_id UUID REFERENCES community_threads(id) ON DELETE CASCADE,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        role TEXT DEFAULT 'MEMBER',
        permissions JSONB DEFAULT '{}'::jsonb,
        joined_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (thread_id, user_id)
      );

      CREATE TABLE IF NOT EXISTS community_messages (
        id UUID PRIMARY KEY,
        thread_id UUID REFERENCES community_threads(id) ON DELETE CASCADE,
        sender_id UUID REFERENCES users(id),
        body TEXT NOT NULL,
        reply_to_message_id UUID REFERENCES community_messages(id),
        is_edited BOOLEAN DEFAULT FALSE,
        edited_at TIMESTAMP,
        is_deleted BOOLEAN DEFAULT FALSE,
        deleted_at TIMESTAMP,
        created_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS seed_flags (
        key TEXT PRIMARY KEY,
        created_at TIMESTAMP DEFAULT NOW()
      );

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

      CREATE TABLE IF NOT EXISTS community_message_reactions (
        id UUID PRIMARY KEY,
        message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        emoji TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (message_id, user_id, emoji)
      );

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

      CREATE TABLE IF NOT EXISTS community_pinned_messages (
        id UUID PRIMARY KEY,
        thread_id UUID REFERENCES community_threads(id) ON DELETE CASCADE,
        message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
        pinned_by UUID REFERENCES users(id),
        pinned_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (thread_id, message_id)
      );

      CREATE TABLE IF NOT EXISTS community_user_presence (
        user_id UUID PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
        status TEXT DEFAULT 'offline',
        last_seen_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS community_message_read_receipts (
        id UUID PRIMARY KEY,
        message_id UUID REFERENCES community_messages(id) ON DELETE CASCADE,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        read_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (message_id, user_id)
      );

      CREATE TABLE IF NOT EXISTS notifications (
        id UUID PRIMARY KEY,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        kind TEXT NOT NULL,
        message TEXT NOT NULL,
        href TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        read_at TIMESTAMP
      );

      ALTER TABLE IF EXISTS users
        ADD COLUMN IF NOT EXISTS avatar_url TEXT;

      ALTER TABLE IF EXISTS users
        ADD COLUMN IF NOT EXISTS user_name TEXT;

      ALTER TABLE IF EXISTS users
        ADD COLUMN IF NOT EXISTS bio TEXT;

      -- Make user_name NOT NULL and add unique constraint
      DO $$
      BEGIN
        -- First, set a default for any NULL values (using email prefix as fallback)
        UPDATE users SET user_name = LOWER(REGEXP_REPLACE(SPLIT_PART(email, '@', 1), '[^a-zA-Z0-9_-]', '', 'g')) || '_' || SUBSTRING(id::text, 1, 8)
        WHERE user_name IS NULL;
        
        -- Update Isaac Wang's username to Isacc1993
        UPDATE users SET user_name = 'Isacc1993' WHERE name = 'Isacc Wang' AND (user_name IS NULL OR user_name != 'Isacc1993');
        
        -- Then add NOT NULL constraint
        ALTER TABLE users ALTER COLUMN user_name SET NOT NULL;
      EXCEPTION
        WHEN OTHERS THEN
          -- If constraint already exists or other error, continue
          NULL;
      END $$;

      CREATE UNIQUE INDEX IF NOT EXISTS idx_users_user_name ON users(user_name);

      -- Backfill new community message columns before creating indexes.
      ALTER TABLE IF EXISTS community_messages
        ADD COLUMN IF NOT EXISTS reply_to_message_id UUID REFERENCES community_messages(id);
      ALTER TABLE IF EXISTS community_messages
        ADD COLUMN IF NOT EXISTS is_edited BOOLEAN DEFAULT FALSE;
      ALTER TABLE IF EXISTS community_messages
        ADD COLUMN IF NOT EXISTS edited_at TIMESTAMP;
      ALTER TABLE IF EXISTS community_messages
        ADD COLUMN IF NOT EXISTS is_deleted BOOLEAN DEFAULT FALSE;
      ALTER TABLE IF EXISTS community_messages
        ADD COLUMN IF NOT EXISTS deleted_at TIMESTAMP;
      ALTER TABLE IF EXISTS community_messages
        ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT NOW();

      ALTER TABLE IF EXISTS applications
        ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'in_review',
        ADD COLUMN IF NOT EXISTS is_reviewed BOOLEAN DEFAULT FALSE,
        ADD COLUMN IF NOT EXISTS reviewed_by UUID REFERENCES users(id),
        ADD COLUMN IF NOT EXISTS reviewed_at TIMESTAMP;
      ALTER TABLE IF EXISTS applications
        ADD COLUMN IF NOT EXISTS company TEXT;

      UPDATE applications SET status = COALESCE(status, 'in_review');
      UPDATE applications SET is_reviewed = FALSE WHERE is_reviewed IS NULL;

      ALTER TABLE IF EXISTS profiles
        ADD COLUMN IF NOT EXISTS resume_template_id UUID REFERENCES resume_templates(id) ON DELETE SET NULL;

      CREATE INDEX IF NOT EXISTS idx_profile_accounts_profile ON profile_accounts(profile_id);
      CREATE INDEX IF NOT EXISTS idx_profile_accounts_email ON profile_accounts(email);
      CREATE INDEX IF NOT EXISTS idx_profile_accounts_status ON profile_accounts(status);
      CREATE INDEX IF NOT EXISTS idx_profile_accounts_last_sync_at ON profile_accounts(last_sync_at);
      CREATE INDEX IF NOT EXISTS idx_profiles_base_info_gin ON profiles USING GIN (base_info);
      CREATE INDEX IF NOT EXISTS idx_profiles_base_resume_gin ON profiles USING GIN (base_resume);
      CREATE INDEX IF NOT EXISTS idx_profiles_created_by ON profiles(created_by);
      CREATE INDEX IF NOT EXISTS idx_profiles_assigned_bidder_id ON profiles(assigned_bidder_id);
      CREATE INDEX IF NOT EXISTS idx_profiles_resume_template_id ON profiles(resume_template_id);
      CREATE INDEX IF NOT EXISTS idx_applications_bidder ON applications(bidder_user_id);
      CREATE INDEX IF NOT EXISTS idx_applications_profile ON applications(profile_id);
      CREATE INDEX IF NOT EXISTS idx_applications_created_at ON applications(created_at);
      CREATE INDEX IF NOT EXISTS idx_applications_domain ON applications(domain);
      CREATE INDEX IF NOT EXISTS idx_applications_reviewed ON applications(is_reviewed, created_at DESC);
      CREATE INDEX IF NOT EXISTS idx_tasks_due ON tasks(due_date);
      CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status);
      CREATE INDEX IF NOT EXISTS idx_tasks_created_by ON tasks(created_by);
      CREATE INDEX IF NOT EXISTS idx_tasks_project ON tasks(project);
      CREATE INDEX IF NOT EXISTS idx_task_assignees_task ON task_assignees(task_id);
      CREATE INDEX IF NOT EXISTS idx_task_assignees_user ON task_assignees(user_id);
      CREATE INDEX IF NOT EXISTS idx_task_assignment_requests_status ON task_assignment_requests(status);
      CREATE INDEX IF NOT EXISTS idx_task_assignment_requests_task ON task_assignment_requests(task_id);
      CREATE INDEX IF NOT EXISTS idx_task_done_requests_status ON task_done_requests(status);
      CREATE INDEX IF NOT EXISTS idx_task_done_requests_task ON task_done_requests(task_id);
      CREATE INDEX IF NOT EXISTS idx_resume_templates_updated ON resume_templates(updated_at);
      CREATE INDEX IF NOT EXISTS idx_resume_templates_name ON resume_templates(name);
      CREATE INDEX IF NOT EXISTS idx_community_members_thread ON community_thread_members(thread_id);
      CREATE INDEX IF NOT EXISTS idx_community_members_user ON community_thread_members(user_id);
      CREATE INDEX IF NOT EXISTS idx_community_members_user_thread ON community_thread_members(user_id, thread_id);
      CREATE INDEX IF NOT EXISTS idx_community_messages_thread ON community_messages(thread_id);
      CREATE INDEX IF NOT EXISTS idx_community_messages_sender ON community_messages(sender_id);
      CREATE INDEX IF NOT EXISTS idx_community_messages_thread_created_at ON community_messages(thread_id, created_at DESC);
      CREATE INDEX IF NOT EXISTS idx_attachments_message ON community_message_attachments(message_id);
      CREATE INDEX IF NOT EXISTS idx_reactions_message ON community_message_reactions(message_id);
      CREATE INDEX IF NOT EXISTS idx_reactions_user ON community_message_reactions(user_id);
      CREATE INDEX IF NOT EXISTS idx_unread_user ON community_unread_messages(user_id);
      CREATE INDEX IF NOT EXISTS idx_unread_thread ON community_unread_messages(thread_id);
      CREATE INDEX IF NOT EXISTS idx_pinned_thread ON community_pinned_messages(thread_id);
      CREATE INDEX IF NOT EXISTS idx_read_receipts_message ON community_message_read_receipts(message_id);
      CREATE INDEX IF NOT EXISTS idx_read_receipts_user ON community_message_read_receipts(user_id);
      CREATE INDEX IF NOT EXISTS idx_notifications_user ON notifications(user_id, created_at DESC);
      CREATE INDEX IF NOT EXISTS idx_notifications_unread ON notifications(user_id) WHERE read_at IS NULL;

      CREATE TABLE IF NOT EXISTS user_oauth_accounts (
        id UUID PRIMARY KEY,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        provider TEXT NOT NULL DEFAULT 'azure-ad',
        provider_account_id TEXT NOT NULL,
        email TEXT NOT NULL,
        display_name TEXT,
        access_token TEXT,
        refresh_token TEXT,
        expires_at BIGINT,
        ext_expires_in INTEGER,
        token_type TEXT,
        scope TEXT,
        id_token TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (user_id, provider, provider_account_id)
      );

      CREATE INDEX IF NOT EXISTS idx_user_oauth_accounts_user ON user_oauth_accounts(user_id);
      CREATE INDEX IF NOT EXISTS idx_user_oauth_accounts_provider ON user_oauth_accounts(provider);
      CREATE INDEX IF NOT EXISTS idx_user_oauth_accounts_user_provider ON user_oauth_accounts(user_id, provider);

      CREATE TABLE IF NOT EXISTS calendar_events (
        id UUID PRIMARY KEY,
        owner_user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        mailbox TEXT,
        mailbox_id UUID REFERENCES user_oauth_accounts(id) ON DELETE SET NULL,
        provider TEXT NOT NULL DEFAULT 'MICROSOFT',
        provider_event_id TEXT NOT NULL,
        title TEXT,
        start_at TEXT NOT NULL,
        end_at TEXT NOT NULL,
        is_all_day BOOLEAN DEFAULT FALSE,
        organizer TEXT,
        location TEXT,
        timezone TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (owner_user_id, provider_event_id)
      );

      CREATE INDEX IF NOT EXISTS idx_calendar_events_owner_mailbox ON calendar_events(owner_user_id, mailbox);
      CREATE INDEX IF NOT EXISTS idx_calendar_events_owner_start ON calendar_events(owner_user_id, start_at);

      CREATE TABLE IF NOT EXISTS daily_reports (
        id UUID PRIMARY KEY,
        user_id UUID REFERENCES users(id) ON DELETE CASCADE,
        report_date DATE NOT NULL,
        status TEXT NOT NULL DEFAULT 'draft',
        content TEXT,
        review_reason TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW(),
        submitted_at TIMESTAMP,
        reviewed_at TIMESTAMP,
        reviewed_by UUID REFERENCES users(id),
        UNIQUE (user_id, report_date)
      );

      CREATE INDEX IF NOT EXISTS idx_daily_reports_user_status ON daily_reports(user_id, status);
      CREATE INDEX IF NOT EXISTS idx_daily_reports_report_date ON daily_reports(report_date);
      CREATE INDEX IF NOT EXISTS idx_daily_reports_reviewed_by ON daily_reports(reviewed_by);

      CREATE TABLE IF NOT EXISTS daily_report_attachments (
        id UUID PRIMARY KEY,
        report_id UUID REFERENCES daily_reports(id) ON DELETE CASCADE,
        file_url TEXT NOT NULL,
        file_name TEXT NOT NULL,
        file_size BIGINT NOT NULL,
        mime_type TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT NOW(),
        UNIQUE (report_id, file_url)
      );

      CREATE INDEX IF NOT EXISTS idx_daily_report_attachments_report ON daily_report_attachments(report_id);
    `);
    
    console.log('✅ [DB] Core tables created/verified');
    console.log('✅ [DB] Performance indexes created/verified (40+ indexes)');
    console.log('\n🗑️  [DB] Cleaning up unused tables...');
    await client.query(`
      DROP TABLE IF EXISTS resumes;
      DROP TABLE IF EXISTS events;
      DROP TABLE IF EXISTS sessions;
      DROP TABLE IF EXISTS llm_settings;
      DROP TABLE IF EXISTS community_channel_roles;
      DROP TABLE IF EXISTS "User";
      DROP TABLE IF EXISTS "Account";
      DROP TABLE IF EXISTS "Session";
      DROP TABLE IF EXISTS "VerificationToken";

      ALTER TABLE IF EXISTS profiles
        ADD COLUMN IF NOT EXISTS base_resume JSONB DEFAULT '{}'::jsonb;
      ALTER TABLE IF EXISTS profiles
        ADD COLUMN IF NOT EXISTS base_additional_bullets JSONB DEFAULT '{}'::jsonb;
      ALTER TABLE IF EXISTS profiles
        ADD COLUMN IF NOT EXISTS assigned_bidder_id UUID REFERENCES users(id);
      ALTER TABLE IF EXISTS profiles
        ADD COLUMN IF NOT EXISTS assigned_by UUID REFERENCES users(id);
      ALTER TABLE IF EXISTS profiles
        ADD COLUMN IF NOT EXISTS assigned_at TIMESTAMP;

      ALTER TABLE IF EXISTS tasks
        ALTER COLUMN due_date DROP NOT NULL;

      DROP TABLE IF EXISTS assignments;
    `);
    console.log('✅ [DB] Unused tables cleaned up (10 tables dropped if existed)');
    
    console.log('\n🔧 [DB] Applying schema updates...');
    await client.query(
      "ALTER TABLE IF EXISTS daily_reports ADD COLUMN IF NOT EXISTS review_reason TEXT;",
    );

    await client.query(
      `
        ALTER TABLE IF EXISTS tasks
          ADD COLUMN IF NOT EXISTS rejected_by UUID REFERENCES users(id),
          ADD COLUMN IF NOT EXISTS rejected_at TIMESTAMP,
          ADD COLUMN IF NOT EXISTS rejection_reason TEXT
      `,
    );

    // Migration: Rename summary to detail and remove href
    await client.query(
      `
        DO $$
        BEGIN
          -- Rename summary to detail if summary exists and detail doesn't
          IF EXISTS (
            SELECT 1 FROM information_schema.columns 
            WHERE table_name = 'tasks' AND column_name = 'summary'
          ) AND NOT EXISTS (
            SELECT 1 FROM information_schema.columns 
            WHERE table_name = 'tasks' AND column_name = 'detail'
          ) THEN
            ALTER TABLE tasks RENAME COLUMN summary TO detail;
          END IF;
          
          -- Drop href column if it exists
          IF EXISTS (
            SELECT 1 FROM information_schema.columns 
            WHERE table_name = 'tasks' AND column_name = 'href'
          ) THEN
            ALTER TABLE tasks DROP COLUMN href;
          END IF;
        END $$;
      `,
    );


    await client.query(
      'CREATE INDEX IF NOT EXISTS idx_community_messages_reply ON community_messages(reply_to_message_id);',
    );

    await client.query(
      `
        UPDATE daily_reports SET status = 'in_review' WHERE status = 'sent';
        UPDATE daily_reports SET status = 'accepted' WHERE status = 'accept';
        UPDATE daily_reports SET status = 'rejected' WHERE status = 'failed';
      `,
    );

    // Migration 019: Add mailbox_id to calendar_events
    await client.query(
      `
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

        -- Add constraint if it doesn't exist
        DO $$
        BEGIN
          IF NOT EXISTS (
            SELECT 1 FROM information_schema.table_constraints 
            WHERE constraint_name = 'fk_calendar_events_mailbox_id' 
            AND table_name = 'calendar_events'
          ) THEN
            ALTER TABLE calendar_events
              ADD CONSTRAINT fk_calendar_events_mailbox_id
              FOREIGN KEY (mailbox_id) REFERENCES user_oauth_accounts(id) ON DELETE SET NULL;
          END IF;
        END $$;

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
      `,
    );

    console.log('✅ [DB] Additional indexes and data migrations applied');

    console.log('\n🌱 [DB] Seeding initial data...');
    
    const seedKey = 'application_phrases_seed';
    const { rows: seedRows } = await client.query<{ key: string }>(
      'SELECT key FROM seed_flags WHERE key = $1',
      [seedKey],
    );
    if (seedRows.length === 0) {
      const { rows: existing } = await client.query<{ id: string }>(
        'SELECT id FROM label_aliases WHERE canonical_key = $1 LIMIT 1',
        [APPLICATION_SUCCESS_KEY],
      );
      if (existing.length === 0) {
        console.log('  📝 [DB] Seeding application success phrases...');
        for (const phrase of APPLICATION_SUCCESS_DEFAULTS) {
          const normalized = normalizeLabelAlias(phrase);
          if (!normalized) continue;
          await client.query(
            `
              INSERT INTO label_aliases (id, canonical_key, alias, normalized_alias)
              VALUES ($1, $2, $3, $4)
              ON CONFLICT (normalized_alias) DO NOTHING
            `,
            [randomUUID(), APPLICATION_SUCCESS_KEY, phrase, normalized],
          );
        }
        console.log('  ✅ [DB] Application phrases seeded');
      }
      await client.query('INSERT INTO seed_flags (key) VALUES ($1)', [seedKey]);
    } else {
      console.log('  ⏭️  [DB] Application phrases already seeded');
    }

    const communitySeedKey = 'community_default_channels_seed';
    const { rows: communitySeedRows } = await client.query<{ key: string }>(
      'SELECT key FROM seed_flags WHERE key = $1',
      [communitySeedKey],
    );
    if (communitySeedRows.length === 0) {
      console.log('  📝 [DB] Creating default community channels...');
      const defaults = [
        {
          name: 'general',
          description: 'Company-wide discussions and daily updates.',
        },
        {
          name: 'announcements',
          description: 'Important notices from the team.',
        },
        {
          name: 'sandbox',
          description: 'Play area to try things out.',
        },
      ];
      for (const channel of defaults) {
        const key = channel.name.trim().toLowerCase();
        if (!key) continue;
        await client.query(
          `
            INSERT INTO community_threads (id, thread_type, name, name_key, description, created_by, is_private)
            VALUES ($1, 'CHANNEL', $2, $3, $4, NULL, FALSE)
            ON CONFLICT (name_key) DO NOTHING
          `,
          [randomUUID(), channel.name.trim(), key, channel.description ?? null],
        );
      }
      console.log('  ✅ [DB] Default community channels created');
      await client.query('INSERT INTO seed_flags (key) VALUES ($1)', [communitySeedKey]);
    } else {
      console.log('  ⏭️  [DB] Community channels already exist');
    }

    const adminSeedKey = 'default_admin_seed';
    const { rows: adminSeedRows } = await client.query<{ key: string }>(
      'SELECT key FROM seed_flags WHERE key = $1',
      [adminSeedKey],
    );
    if (adminSeedRows.length === 0) {
      console.log('  📝 [DB] Creating default admin user...');
      await client.query(
        `
          INSERT INTO users (id, email, user_name, name, role, password, is_active, created_at)
          VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
          ON CONFLICT (email) DO UPDATE SET user_name = EXCLUDED.user_name
        `,
        [
          'b288230a-590d-4030-85d4-c72dfe7f5ae6',
          'wrenikey.dev@gmail.com',
          'Isacc1993',
          'Isacc Wang',
          'ADMIN',
          '$2a$08$TgVEMWkdL0OQnqpAbXGoJe/sUpU9iRaTNvKBN25pJAr2TWQoZ4yC2',
          true,
          '2025-12-12 00:52:11.345273',
        ],
      );
      console.log('  ✅ [DB] Default admin user created');
      await client.query('INSERT INTO seed_flags (key) VALUES ($1)', [adminSeedKey]);
    } else {
      console.log('  ⏭️  [DB] Default admin user already exists');
    }

    console.log('\n✅ [DB] Database initialization completed successfully!\n');
  } catch (error) {
    console.error('\n❌ [DB] Database initialization failed:');
    console.error(error);
    throw error;
  } finally {
    client.release();
  }
}

export async function findUserByEmail(email: string) {
  const { rows } = await pool.query<User>(
    'SELECT id, email, user_name AS "userName", name, avatar_url AS "avatarUrl", bio, role, is_active as "isActive", password FROM users WHERE email = $1',
    [email],
  );
  return rows[0];
}

export async function findUserByUserName(userName: string) {
  const { rows } = await pool.query<User>(
    'SELECT id, email, user_name AS "userName", name, avatar_url AS "avatarUrl", bio, role, is_active as "isActive", password FROM users WHERE user_name = $1',
    [userName],
  );
  return rows[0];
}

export async function findUserById(id: string) {
  const { rows } = await pool.query<User>(
    'SELECT id, email, user_name AS "userName", name, avatar_url AS "avatarUrl", bio, role, is_active as "isActive", password FROM users WHERE id = $1',
    [id],
  );
  return rows[0];
}

export async function listActiveUserIds(): Promise<string[]> {
  const { rows } = await pool.query<{ id: string }>(
    'SELECT id FROM users WHERE is_active = TRUE',
  );
  return rows.map((row) => row.id);
}

export async function listActiveUserIdsByRole(roles: string[]): Promise<string[]> {
  if (!roles.length) return [];
  const { rows } = await pool.query<{ id: string }>(
    'SELECT id FROM users WHERE is_active = TRUE AND role = ANY($1)',
    [roles],
  );
  return rows.map((row) => row.id);
}

export async function insertUser(user: User) {
  const avatarUrl =
    user.avatarUrl && user.avatarUrl.toLowerCase() !== 'nope' ? user.avatarUrl : null;
  await pool.query(
    `
      INSERT INTO users (id, email, user_name, name, avatar_url, role, password, is_active)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
      ON CONFLICT (email) DO UPDATE SET
        user_name = EXCLUDED.user_name,
        name = EXCLUDED.name,
        avatar_url = EXCLUDED.avatar_url,
        role = EXCLUDED.role,
        password = EXCLUDED.password,
        is_active = EXCLUDED.is_active
    `,
    [
      user.id,
      user.email,
      user.userName,
      user.name,
      avatarUrl,
      user.role,
      user.password ?? 'demo',
      user.isActive ?? true,
    ],
  );
}

export async function updateUserUserName(userId: string, userName: string) {
  await pool.query('UPDATE users SET user_name = $1 WHERE id = $2', [userName, userId]);
}

export async function updateUserAvatar(userId: string, avatarUrl: string | null) {
  const normalized =
    avatarUrl && avatarUrl.toLowerCase() !== 'nope' ? avatarUrl : null;
  await pool.query('UPDATE users SET avatar_url = $1 WHERE id = $2', [
    normalized,
    userId,
  ]);
}

export async function updateUserNameAndEmail(userId: string, name: string, email: string, bio?: string | null) {
  if (bio !== undefined) {
    await pool.query('UPDATE users SET name = $1, email = $2, bio = $3 WHERE id = $4', [name, email, bio, userId]);
  } else {
    await pool.query('UPDATE users SET name = $1, email = $2 WHERE id = $3', [name, email, userId]);
  }
}

export async function updateUserBio(userId: string, bio: string | null) {
  await pool.query('UPDATE users SET bio = $1 WHERE id = $2', [bio, userId]);
}

export async function updateUserPassword(userId: string, hashedPassword: string) {
  await pool.query('UPDATE users SET password = $1 WHERE id = $2', [hashedPassword, userId]);
}

export async function insertProfile(profile: {
  id: string;
  displayName: string;
  baseInfo: Record<string, unknown>;
  baseResume?: Record<string, unknown>;
  baseAdditionalBullets?: Record<string, number>;
  resumeTemplateId?: string | null;
  createdBy?: string;
  createdAt?: string;
  updatedAt?: string;
}) {
  await pool.query(
    `
      INSERT INTO profiles (id, display_name, base_info, base_resume, base_additional_bullets, resume_template_id, created_by, created_at, updated_at)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
      ON CONFLICT (id) DO NOTHING
    `,
    [
      profile.id,
      profile.displayName,
      JSON.stringify(profile.baseInfo ?? {}),
      JSON.stringify(profile.baseResume ?? {}),
      JSON.stringify(profile.baseAdditionalBullets ?? {}),
      profile.resumeTemplateId ?? null,
      profile.createdBy ?? null,
      profile.createdAt ?? new Date().toISOString(),
      profile.updatedAt ?? new Date().toISOString(),
    ],
  );
}

export async function insertApplication(record: ApplicationRecord) {
  await pool.query(
    `
      INSERT INTO applications (
        id,
        session_id,
        bidder_user_id,
        profile_id,
        resume_id,
        url,
        domain,
        company,
        status,
        created_at,
        is_reviewed,
        reviewed_by,
        reviewed_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13)
      ON CONFLICT (session_id) DO NOTHING
    `,
    [
      record.id,
      record.sessionId,
      record.bidderUserId,
      record.profileId,
      record.resumeId ?? null,
      record.url,
      record.domain ?? null,
      record.company ?? null,
      record.createdAt,
      record.status ?? "in_review",
      record.isReviewed ?? false,
      record.reviewedBy ?? null,
      record.reviewedAt ?? null,
    ],
  );
}

export async function insertApplicationWithSummary(
  record: ApplicationRecord,
): Promise<ApplicationSummary | undefined> {
  await insertApplication(record);
  const { rows } = await pool.query<ApplicationSummary>(
    `
      ${APPLICATION_SUMMARY_SELECT}
      WHERE a.id = $1
    `,
    [record.id],
  );
  return rows[0];
}

export async function listProfiles(): Promise<Profile[]> {
  const { rows } = await pool.query<Profile>(
    `
      SELECT
        p.id,
        p.display_name AS "displayName",
        p.base_info AS "baseInfo",
        p.base_resume AS "baseResume",
        p.base_additional_bullets AS "baseAdditionalBullets",
        p.resume_template_id AS "resumeTemplateId",
        rt.name AS "resumeTemplateName",
        p.created_by AS "createdBy",
        p.assigned_bidder_id AS "assignedBidderId",
        p.assigned_by AS "assignedBy",
        p.assigned_at AS "assignedAt",
        p.created_at AS "createdAt",
        p.updated_at AS "updatedAt"
      FROM profiles p
      LEFT JOIN resume_templates rt ON rt.id = p.resume_template_id
      ORDER BY p.created_at DESC
    `,
  );
  return rows;
}

export async function listProfilesForBidder(bidderUserId: string): Promise<Profile[]> {
  const { rows } = await pool.query<Profile>(
    `
      SELECT
        p.id,
        p.display_name AS "displayName",
        p.base_info AS "baseInfo",
        p.base_resume AS "baseResume",
        p.base_additional_bullets AS "baseAdditionalBullets",
        p.resume_template_id AS "resumeTemplateId",
        rt.name AS "resumeTemplateName",
        p.created_by AS "createdBy",
        p.assigned_bidder_id AS "assignedBidderId",
        p.assigned_by AS "assignedBy",
        p.assigned_at AS "assignedAt",
        p.created_at AS "createdAt",
        p.updated_at AS "updatedAt"
      FROM profiles p
      LEFT JOIN resume_templates rt ON rt.id = p.resume_template_id
      WHERE p.assigned_bidder_id = $1
      ORDER BY p.created_at DESC
    `,
    [bidderUserId],
  );
  return rows;
}

export async function findProfileById(id: string): Promise<Profile | undefined> {
  const { rows } = await pool.query<Profile>(
    `
      SELECT
        p.id,
        p.display_name AS "displayName",
        p.base_info AS "baseInfo",
        p.base_resume AS "baseResume",
        p.base_additional_bullets AS "baseAdditionalBullets",
        p.resume_template_id AS "resumeTemplateId",
        rt.name AS "resumeTemplateName",
        p.created_by AS "createdBy",
        p.assigned_bidder_id AS "assignedBidderId",
        p.assigned_by AS "assignedBy",
        p.assigned_at AS "assignedAt",
        p.created_at AS "createdAt",
        p.updated_at AS "updatedAt"
      FROM profiles p
      LEFT JOIN resume_templates rt ON rt.id = p.resume_template_id
      WHERE p.id = $1
    `,
    [id],
  );
  return rows[0];
}

export async function listProfileAccountsForUser(
  actor: User,
  profileId?: string,
): Promise<ProfileAccountWithProfile[]> {
  const { rows } = await pool.query<ProfileAccountWithProfile>(
    `
      SELECT
        pa.id,
        pa.profile_id AS "profileId",
        pa.provider,
        pa.email,
        pa.display_name AS "displayName",
        pa.timezone,
        pa.status,
        pa.last_sync_at AS "lastSyncAt",
        pa.created_at AS "createdAt",
        pa.updated_at AS "updatedAt",
        p.display_name AS "profileDisplayName",
        p.assigned_bidder_id AS "profileAssignedBidderId"
      FROM profile_accounts pa
      JOIN profiles p ON p.id = pa.profile_id
      WHERE
        ($1 = 'ADMIN' OR $1 = 'MANAGER' OR p.assigned_bidder_id = $2)
        AND ($3::uuid IS NULL OR pa.profile_id = $3)
      ORDER BY pa.updated_at DESC, pa.created_at DESC
    `,
    [actor.role, actor.id, profileId ?? null],
  );
  return rows;
}

export async function findProfileAccountById(id: string): Promise<ProfileAccountWithProfile | undefined> {
  const { rows } = await pool.query<ProfileAccountWithProfile>(
    `
      SELECT
        pa.id,
        pa.profile_id AS "profileId",
        pa.provider,
        pa.email,
        pa.display_name AS "displayName",
        pa.timezone,
        pa.status,
        pa.last_sync_at AS "lastSyncAt",
        pa.created_at AS "createdAt",
        pa.updated_at AS "updatedAt",
        p.display_name AS "profileDisplayName",
        p.assigned_bidder_id AS "profileAssignedBidderId"
      FROM profile_accounts pa
      JOIN profiles p ON p.id = pa.profile_id
      WHERE pa.id = $1
      LIMIT 1
    `,
    [id],
  );
  return rows[0];
}

export async function upsertProfileAccount(account: {
  id: string;
  profileId: string;
  provider?: string;
  email: string;
  displayName?: string | null;
  timezone?: string | null;
  status?: string | null;
  lastSyncAt?: string | null;
}): Promise<ProfileAccount> {
  const { rows } = await pool.query<ProfileAccount>(
    `
      INSERT INTO profile_accounts (id, profile_id, provider, email, display_name, timezone, status, last_sync_at)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
      ON CONFLICT (profile_id, email) DO UPDATE
        SET provider = EXCLUDED.provider,
            display_name = EXCLUDED.display_name,
            timezone = EXCLUDED.timezone,
            status = EXCLUDED.status,
            last_sync_at = COALESCE(EXCLUDED.last_sync_at, profile_accounts.last_sync_at),
            updated_at = NOW()
      RETURNING
        id,
        profile_id AS "profileId",
        provider,
        email,
        display_name AS "displayName",
        timezone,
        status,
        last_sync_at AS "lastSyncAt",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
    `,
    [
      account.id,
      account.profileId,
      account.provider ?? 'MICROSOFT',
      account.email,
      account.displayName ?? null,
      account.timezone ?? 'UTC',
      account.status ?? 'ACTIVE',
      account.lastSyncAt ?? null,
    ],
  );
  return rows[0];
}

export async function touchProfileAccount(id: string, lastSyncAt?: string) {
  await pool.query(
    `
      UPDATE profile_accounts
      SET last_sync_at = $2,
          updated_at = NOW()
      WHERE id = $1
    `,
    [id, lastSyncAt ?? new Date().toISOString()],
  );
}


/**
 * Get mailbox_id (user_oauth_accounts.id) from email address
 * For non-admin users, also checks admin's accounts
 */
export async function getMailboxIdFromEmail(
  ownerUserId: string,
  email: string,
  isAdmin: boolean = false,
): Promise<string | null> {
  // First try to find in owner's accounts
  const { rows: ownerRows } = await pool.query<{ id: string }>(
    `
      SELECT id
      FROM user_oauth_accounts
      WHERE user_id = $1 AND LOWER(email) = LOWER($2)
      LIMIT 1
    `,
    [ownerUserId, email],
  );

  if (ownerRows.length > 0) {
    return ownerRows[0].id;
  }

  // For non-admin users, also check admin's accounts
  if (!isAdmin) {
    const { rows: adminRows } = await pool.query<{ id: string; user_id: string }>(
      `
        SELECT uoa.id, uoa.user_id
        FROM user_oauth_accounts uoa
        INNER JOIN users u ON u.id = uoa.user_id
        WHERE u.role = 'ADMIN' AND LOWER(uoa.email) = LOWER($1)
        ORDER BY u.created_at ASC
        LIMIT 1
      `,
      [email],
    );

    if (adminRows.length > 0) {
      return adminRows[0].id;
    }
  }

  return null;
}

export async function listCalendarEventsForOwner(
  ownerUserId: string,
  mailboxIds?: string[],
  range?: { start?: string | null; end?: string | null },
): Promise<StoredCalendarEvent[]> {
  const { rows } = await pool.query<StoredCalendarEvent & { mailbox_email: string | null }>(
    `
      SELECT
        ce.provider_event_id AS id,
        ce.mailbox_id AS "mailboxId",
        COALESCE(uoa.email, ce.mailbox) AS mailbox,
        ce.title,
        ce.start_at AS "start",
        ce.end_at AS "end",
        ce.is_all_day AS "isAllDay",
        ce.organizer,
        ce.location,
        ce.timezone
      FROM calendar_events ce
      LEFT JOIN user_oauth_accounts uoa ON ce.mailbox_id = uoa.id
      WHERE ce.owner_user_id = $1
        AND ($2::uuid[] IS NULL OR ce.mailbox_id = ANY($2))
        AND ($3::text IS NULL OR ce.end_at >= $3)
        AND ($4::text IS NULL OR ce.start_at <= $4)
      ORDER BY ce.start_at ASC
    `,
    [
      ownerUserId,
      mailboxIds && mailboxIds.length ? mailboxIds : null,
      range?.start ?? null,
      range?.end ?? null,
    ],
  );
  return rows.map(row => ({
    ...row,
    mailbox: row.mailbox || undefined,  // Convert null to undefined for compatibility
  }));
}

export async function replaceCalendarEvents(params: {
  ownerUserId: string;
  mailboxIds: string[];
  timezone?: string | null;
  events: CalendarEventInput[];
}): Promise<StoredCalendarEvent[]> {
  const { ownerUserId, mailboxIds, timezone, events } = params;
  console.log('[replaceCalendarEvents] Called with:', {
    ownerUserId,
    mailboxIds: mailboxIds.length,
    eventCount: events.length,
    timezone
  });
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    if (mailboxIds.length) {
      console.log('[replaceCalendarEvents] Deleting existing events for mailboxIds:', mailboxIds);
      await client.query(
        'DELETE FROM calendar_events WHERE owner_user_id = $1 AND mailbox_id = ANY($2)',
        [ownerUserId, mailboxIds],
      );
    }
    if (events.length) {
      console.log('[replaceCalendarEvents] Inserting', events.length, 'events');
      const values: string[] = [];
      const args: Array<string | boolean | null> = [];
      let idx = 1;
      for (const event of events) {
        if (!event.mailboxId) continue;
        values.push(
          `($${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++})`,
        );
        args.push(
          randomUUID(),
          ownerUserId,
          event.mailboxId,
          'MICROSOFT',
          event.id,
          event.title ?? null,
          event.start,
          event.end,
          event.isAllDay ?? false,
          event.organizer ?? null,
          event.location ?? null,
          timezone ?? null,
        );
      }
      if (values.length) {
        console.log('[replaceCalendarEvents] Executing INSERT with', values.length, 'events');
        const result = await client.query(
          `
            INSERT INTO calendar_events (
              id,
              owner_user_id,
              mailbox_id,
              provider,
              provider_event_id,
              title,
              start_at,
              end_at,
              is_all_day,
              organizer,
              location,
              timezone
            )
            VALUES ${values.join(', ')}
          `,
          args,
        );
        console.log('[replaceCalendarEvents] INSERT successful, rows affected:', result.rowCount);
      } else {
        console.warn('[replaceCalendarEvents] No values to insert - all events were filtered out');
      }
    }
    await client.query('COMMIT');
    console.log('[replaceCalendarEvents] Transaction COMMITTED');
  } catch (err) {
    console.error('[replaceCalendarEvents] ERROR in transaction:', err);
    await client.query('ROLLBACK');
    throw err;
  } finally {
    client.release();
  }

  console.log('[replaceCalendarEvents] Fetching saved events...');
  const saved = await listCalendarEventsForOwner(ownerUserId, mailboxIds);
  console.log('[replaceCalendarEvents] Retrieved', saved.length, 'events from database');
  return saved;
}

export async function upsertUserOAuthAccount(params: {
  id: string;
  userId: string;
  provider: string;
  providerAccountId: string;
  email: string;
  displayName?: string | null;
  accessToken?: string | null;
  refreshToken?: string | null;
  expiresAt?: number | null;
  extExpiresIn?: number | null;
  tokenType?: string | null;
  scope?: string | null;
  idToken?: string | null;
}): Promise<UserOAuthAccount> {
  const {
    id,
    userId,
    provider,
    providerAccountId,
    email,
    displayName,
    accessToken,
    refreshToken,
    expiresAt,
    extExpiresIn,
    tokenType,
    scope,
    idToken,
  } = params;
  const { rows } = await pool.query<UserOAuthAccount>(
    `
      INSERT INTO user_oauth_accounts (
        id,
        user_id,
        provider,
        provider_account_id,
        email,
        display_name,
        access_token,
        refresh_token,
        expires_at,
        ext_expires_in,
        token_type,
        scope,
        id_token,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, NOW())
      ON CONFLICT (user_id, provider, provider_account_id)
      DO UPDATE SET
        email = EXCLUDED.email,
        display_name = EXCLUDED.display_name,
        access_token = EXCLUDED.access_token,
        refresh_token = COALESCE(EXCLUDED.refresh_token, user_oauth_accounts.refresh_token),
        expires_at = EXCLUDED.expires_at,
        ext_expires_in = EXCLUDED.ext_expires_in,
        token_type = EXCLUDED.token_type,
        scope = EXCLUDED.scope,
        id_token = EXCLUDED.id_token,
        updated_at = NOW()
      RETURNING
        id,
        user_id AS "userId",
        provider,
        provider_account_id AS "providerAccountId",
        email,
        display_name AS "displayName",
        access_token AS "accessToken",
        refresh_token AS "refreshToken",
        expires_at AS "expiresAt",
        ext_expires_in AS "extExpiresIn",
        token_type AS "tokenType",
        scope,
        id_token AS "idToken",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
    `,
    [
      id,
      userId,
      provider,
      providerAccountId,
      email,
      displayName ?? null,
      accessToken ?? null,
      refreshToken ?? null,
      expiresAt ?? null,
      extExpiresIn ?? null,
      tokenType ?? null,
      scope ?? null,
      idToken ?? null,
    ],
  );
  return rows[0];
}

export async function findUserOAuthAccountByProviderId(
  userId: string,
  providerAccountId: string,
  provider: string = 'azure-ad',
): Promise<UserOAuthAccount | null> {
  const { rows } = await pool.query<UserOAuthAccount>(
    `
      SELECT
        id,
        user_id AS "userId",
        provider,
        provider_account_id AS "providerAccountId",
        email,
        display_name AS "displayName",
        access_token AS "accessToken",
        refresh_token AS "refreshToken",
        expires_at AS "expiresAt",
        ext_expires_in AS "extExpiresIn",
        token_type AS "tokenType",
        scope,
        id_token AS "idToken",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
      FROM user_oauth_accounts
      WHERE user_id = $1 AND provider_account_id = $2 AND provider = $3
      LIMIT 1
    `,
    [userId, providerAccountId, provider],
  );
  return rows[0] ?? null;
}

export async function listUserOAuthAccounts(
  userId: string,
  provider?: string,
): Promise<UserOAuthAccount[]> {
  const { rows } = await pool.query<UserOAuthAccount>(
    `
      SELECT
        id,
        user_id AS "userId",
        provider,
        provider_account_id AS "providerAccountId",
        email,
        display_name AS "displayName",
        access_token AS "accessToken",
        refresh_token AS "refreshToken",
        expires_at AS "expiresAt",
        ext_expires_in AS "extExpiresIn",
        token_type AS "tokenType",
        scope,
        id_token AS "idToken",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
      FROM user_oauth_accounts
      WHERE user_id = $1
        AND ($2::text IS NULL OR provider = $2)
      ORDER BY created_at DESC
    `,
    [userId, provider ?? null],
  );
  return rows;
}

export async function findUserOAuthAccountById(
  accountId: string,
): Promise<UserOAuthAccount | null> {
  const { rows } = await pool.query<UserOAuthAccount>(
    `
      SELECT
        id,
        user_id AS "userId",
        provider,
        provider_account_id AS "providerAccountId",
        email,
        display_name AS "displayName",
        access_token AS "accessToken",
        refresh_token AS "refreshToken",
        expires_at AS "expiresAt",
        ext_expires_in AS "extExpiresIn",
        token_type AS "tokenType",
        scope,
        id_token AS "idToken",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
      FROM user_oauth_accounts
      WHERE id = $1
      LIMIT 1
    `,
    [accountId],
  );
  return rows[0] ?? null;
}

export async function deleteUserOAuthAccount(
  userId: string,
  accountId: string,
): Promise<boolean> {
  const { rowCount } = await pool.query(
    `
      DELETE FROM user_oauth_accounts
      WHERE id = $1 AND user_id = $2
    `,
    [accountId, userId],
  );
  return (rowCount ?? 0) > 0;
}

export async function updateUserOAuthAccountTokens(
  accountId: string,
  tokens: {
    accessToken?: string | null;
    refreshToken?: string | null;
    expiresAt?: number | null;
    extExpiresIn?: number | null;
    idToken?: string | null;
  },
): Promise<UserOAuthAccount | null> {
  const { accessToken, refreshToken, expiresAt, extExpiresIn, idToken } = tokens;
  const { rows } = await pool.query<UserOAuthAccount>(
    `
      UPDATE user_oauth_accounts
      SET
        access_token = COALESCE($2, access_token),
        refresh_token = COALESCE($3, refresh_token),
        expires_at = COALESCE($4, expires_at),
        ext_expires_in = COALESCE($5, ext_expires_in),
        id_token = COALESCE($6, id_token),
        updated_at = NOW()
      WHERE id = $1
      RETURNING
        id,
        user_id AS "userId",
        provider,
        provider_account_id AS "providerAccountId",
        email,
        display_name AS "displayName",
        access_token AS "accessToken",
        refresh_token AS "refreshToken",
        expires_at AS "expiresAt",
        ext_expires_in AS "extExpiresIn",
        token_type AS "tokenType",
        scope,
        id_token AS "idToken",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
    `,
    [accountId, accessToken ?? null, refreshToken ?? null, expiresAt ?? null, extExpiresIn ?? null, idToken ?? null],
  );
  return rows[0] ?? null;
}

export async function listDailyReportsForUser(
  userId: string,
  range: { start?: string | null; end?: string | null } = {},
): Promise<DailyReport[]> {
  const params: Array<string | null> = [userId];
  let query = `
    SELECT
      id,
      user_id AS "userId",
      report_date::text AS "reportDate",
      status,
      content,
      review_reason AS "reviewReason",
      created_at AS "createdAt",
      updated_at AS "updatedAt",
      submitted_at AS "submittedAt",
      reviewed_at AS "reviewedAt",
      reviewed_by AS "reviewedBy"
    FROM daily_reports
    WHERE user_id = $1
  `;
  if (range.start) {
    params.push(range.start);
    query += ` AND report_date >= $${params.length}`;
  }
  if (range.end) {
    params.push(range.end);
    query += ` AND report_date <= $${params.length}`;
  }
  query += ' ORDER BY report_date ASC';
  const { rows } = await pool.query<DailyReport>(query, params);
  return rows;
}

export async function countDailyReportsInReview(
  since?: string | null,
): Promise<number> {
  const params: Array<string> = [];
  let query = `
    SELECT COUNT(*)::int AS count
    FROM daily_reports
    WHERE status = 'in_review'
  `;
  if (since) {
    params.push(since);
    query += ` AND submitted_at IS NOT NULL AND submitted_at >= $1`;
  }
  const { rows } = await pool.query<{ count: number }>(query, params);
  return rows[0]?.count ?? 0;
}

export async function countReviewedDailyReportsForUser(
  userId: string,
  since?: string | null,
): Promise<number> {
  const params: Array<string> = [userId];
  let query = `
    SELECT COUNT(*)::int AS count
    FROM daily_reports
    WHERE user_id = $1
      AND status IN ('accepted', 'rejected')
      AND reviewed_at IS NOT NULL
  `;
  if (since) {
    params.push(since);
    query += ` AND reviewed_at > $${params.length}`;
  }
  const { rows } = await pool.query<{ count: number }>(query, params);
  return rows[0]?.count ?? 0;
}

export async function insertNotifications(
  entries: Array<{
    userId: string;
    kind: string;
    message: string;
    href?: string | null;
    createdAt?: string | null;
  }>,
) {
  if (!entries.length) return;
  const values: string[] = [];
  const params: Array<string | null> = [];
  let idx = 1;
  entries.forEach((entry) => {
    values.push(
      `($${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++})`,
    );
    params.push(
      randomUUID(),
      entry.userId,
      entry.kind,
      entry.message,
      entry.href ?? null,
      entry.createdAt ?? new Date().toISOString(),
    );
  });
  await pool.query(
    `
      INSERT INTO notifications (id, user_id, kind, message, href, created_at)
      VALUES ${values.join(', ')}
    `,
    params,
  );
}

export async function listNotificationsForUser(
  userId: string,
  options: { unreadOnly?: boolean; limit?: number } = {},
): Promise<Notification[]> {
  const params: Array<string | number> = [userId];
  let query = `
    SELECT
      id,
      user_id AS "userId",
      kind,
      message,
      href,
      created_at AS "createdAt",
      read_at AS "readAt"
    FROM notifications
    WHERE user_id = $1
  `;
  if (options.unreadOnly) {
    query += ' AND read_at IS NULL';
  }
  query += ' ORDER BY created_at DESC';
  if (options.limit) {
    params.push(options.limit);
    query += ` LIMIT $${params.length}`;
  }
  const { rows } = await pool.query<Notification>(query, params);
  return rows;
}

export async function markNotificationsRead(
  userId: string,
  ids?: string[],
): Promise<void> {
  if (ids && ids.length > 0) {
    await pool.query(
      `
        UPDATE notifications
        SET read_at = NOW()
        WHERE user_id = $1 AND id = ANY($2) AND read_at IS NULL
      `,
      [userId, ids],
    );
    return;
  }
  await pool.query(
    `
      UPDATE notifications
      SET read_at = NOW()
      WHERE user_id = $1 AND read_at IS NULL
    `,
    [userId],
  );
}

export async function countUnreadNotifications(userId: string): Promise<number> {
  const { rows } = await pool.query<{ count: number }>(
    `
      SELECT COUNT(*)::int AS count
      FROM notifications
      WHERE user_id = $1 AND read_at IS NULL
    `,
    [userId],
  );
  return rows[0]?.count ?? 0;
}

export async function listReviewedDailyReportsForUser(
  userId: string,
  since?: string | null,
): Promise<Array<{ id: string; reportDate: string; status: DailyReport['status']; reviewedAt: string }>> {
  const params: Array<string> = [userId];
  let query = `
    SELECT
      id,
      report_date::text AS "reportDate",
      status,
      reviewed_at AS "reviewedAt"
    FROM daily_reports
    WHERE user_id = $1
      AND status IN ('accepted', 'rejected')
      AND reviewed_at IS NOT NULL
  `;
  if (since) {
    params.push(since);
    query += ` AND reviewed_at > $${params.length}`;
  }
  query += ' ORDER BY reviewed_at DESC';
  const { rows } = await pool.query<{
    id: string;
    reportDate: string;
    status: DailyReport['status'];
    reviewedAt: string;
  }>(query, params);
  return rows;
}

export async function listInReviewReportsWithUsers(): Promise<
  Array<{ id: string; reportDate: string; submittedAt: string | null; updatedAt: string; userName: string }>
> {
  const { rows } = await pool.query<{
    id: string;
    reportDate: string;
    submittedAt: string | null;
    updatedAt: string;
    userName: string;
  }>(
    `
      SELECT
        r.id,
        r.report_date::text AS "reportDate",
        r.submitted_at AS "submittedAt",
        r.updated_at AS "updatedAt",
        u.name AS "userName"
      FROM daily_reports r
      JOIN users u ON u.id = r.user_id
      WHERE r.status = 'in_review'
      ORDER BY r.submitted_at DESC NULLS LAST, r.updated_at DESC
    `,
  );
  return rows;
}

export async function listUnreadCommunityNotifications(userId: string): Promise<
  Array<{
    threadId: string;
    threadType: string;
    threadName: string | null;
    unreadCount: number;
    messageId: string | null;
    messageBody: string | null;
    messageCreatedAt: string | null;
    senderName: string | null;
  }>
> {
  const { rows } = await pool.query<{
    threadId: string;
    threadType: string;
    threadName: string | null;
    unreadCount: number;
    messageId: string | null;
    messageBody: string | null;
    messageCreatedAt: string | null;
    senderName: string | null;
  }>(
    `
      SELECT
        u.thread_id AS "threadId",
        t.thread_type AS "threadType",
        t.name AS "threadName",
        u.unread_count AS "unreadCount",
        m.id AS "messageId",
        m.body AS "messageBody",
        m.created_at AS "messageCreatedAt",
        sender.user_name AS "senderName"
      FROM community_unread_messages u
      JOIN community_threads t ON t.id = u.thread_id
      LEFT JOIN LATERAL (
        SELECT id, body, created_at, sender_id
        FROM community_messages
        WHERE thread_id = u.thread_id AND created_at > u.last_read_at
        ORDER BY created_at DESC
        LIMIT 1
      ) m ON TRUE
      LEFT JOIN users sender ON sender.id = m.sender_id
      WHERE u.user_id = $1 AND u.unread_count > 0
      ORDER BY COALESCE(m.created_at, u.updated_at) DESC
    `,
    [userId],
  );
  return rows;
}

export async function listInReviewReports(range: {
  start: string;
  end: string;
}): Promise<
  Array<{
    id: string;
    userId: string;
    reportDate: string;
    updatedAt: string;
  }>
> {
  const { rows } = await pool.query<{
    id: string;
    userId: string;
    reportDate: string;
    updatedAt: string;
  }>(
    `
      SELECT
        id,
        user_id AS "userId",
        report_date::text AS "reportDate",
        updated_at AS "updatedAt"
      FROM daily_reports
      WHERE status = 'in_review'
        AND report_date BETWEEN $1 AND $2
      ORDER BY report_date ASC, updated_at DESC
    `,
    [range.start, range.end],
  );
  return rows;
}

export async function listAcceptedCountsByDate(range: {
  start: string;
  end: string;
}): Promise<Array<{ reportDate: string; count: number }>> {
  const { rows } = await pool.query<{ reportDate: string; count: number }>(
    `
      SELECT
        report_date::text AS "reportDate",
        COUNT(*)::int AS count
      FROM daily_reports
      WHERE status = 'accepted'
        AND report_date BETWEEN $1 AND $2
      GROUP BY report_date
      ORDER BY report_date ASC
    `,
    [range.start, range.end],
  );
  return rows;
}

export async function listDailyReportsByDate(
  reportDate: string,
): Promise<DailyReportWithUser[]> {
  const { rows } = await pool.query<DailyReportWithUser>(
    `
      SELECT
        r.id,
        r.user_id AS "userId",
        r.report_date::text AS "reportDate",
        r.status,
        r.content,
        r.review_reason AS "reviewReason",
        r.created_at AS "createdAt",
        r.updated_at AS "updatedAt",
        r.submitted_at AS "submittedAt",
        r.reviewed_at AS "reviewedAt",
        r.reviewed_by AS "reviewedBy",
        u.user_name AS "userName",
        u.email AS "userEmail",
        u.avatar_url AS "userAvatarUrl",
        u.role AS "userRole"
      FROM daily_reports r
      JOIN users u ON u.id = r.user_id
      WHERE r.report_date = $1
      ORDER BY u.user_name ASC
    `,
    [reportDate],
  );
  return rows;
}

export async function listDailyReportAttachments(
  reportId: string,
): Promise<DailyReportAttachment[]> {
  const { rows } = await pool.query<DailyReportAttachment>(
    `
      SELECT
        id,
        report_id AS "reportId",
        file_url AS "fileUrl",
        file_name AS "fileName",
        file_size AS "fileSize",
        mime_type AS "mimeType",
        created_at AS "createdAt"
      FROM daily_report_attachments
      WHERE report_id = $1
      ORDER BY created_at ASC
    `,
    [reportId],
  );
  return rows;
}

export async function insertDailyReportAttachments(
  reportId: string,
  attachments: Array<{
    fileUrl: string;
    fileName: string;
    fileSize: number;
    mimeType: string;
  }>,
): Promise<DailyReportAttachment[]> {
  if (!attachments.length) return [];
  const values: string[] = [];
  const params: Array<string | number> = [];
  let idx = 1;
  attachments.forEach((attachment) => {
    values.push(`($${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, $${idx++}, NOW())`);
    params.push(
      randomUUID(),
      reportId,
      attachment.fileUrl,
      attachment.fileName,
      attachment.fileSize,
      attachment.mimeType,
    );
  });
  const { rows } = await pool.query<DailyReportAttachment>(
    `
      INSERT INTO daily_report_attachments (
        id,
        report_id,
        file_url,
        file_name,
        file_size,
        mime_type,
        created_at
      )
      VALUES ${values.join(', ')}
      ON CONFLICT (report_id, file_url) DO NOTHING
      RETURNING
        id,
        report_id AS "reportId",
        file_url AS "fileUrl",
        file_name AS "fileName",
        file_size AS "fileSize",
        mime_type AS "mimeType",
        created_at AS "createdAt"
    `,
    params,
  );
  return rows;
}

export async function findDailyReportByUserAndDate(
  userId: string,
  reportDate: string,
): Promise<DailyReport | undefined> {
  const { rows } = await pool.query<DailyReport>(
    `
      SELECT
        id,
        user_id AS "userId",
        report_date::text AS "reportDate",
        status,
        content,
        review_reason AS "reviewReason",
        created_at AS "createdAt",
        updated_at AS "updatedAt",
        submitted_at AS "submittedAt",
        reviewed_at AS "reviewedAt",
        reviewed_by AS "reviewedBy"
      FROM daily_reports
      WHERE user_id = $1 AND report_date = $2
      LIMIT 1
    `,
    [userId, reportDate],
  );
  return rows[0];
}

export async function findDailyReportById(id: string): Promise<DailyReport | undefined> {
  const { rows } = await pool.query<DailyReport>(
    `
      SELECT
        id,
        user_id AS "userId",
        report_date::text AS "reportDate",
        status,
        content,
        review_reason AS "reviewReason",
        created_at AS "createdAt",
        updated_at AS "updatedAt",
        submitted_at AS "submittedAt",
        reviewed_at AS "reviewedAt",
        reviewed_by AS "reviewedBy"
      FROM daily_reports
      WHERE id = $1
      LIMIT 1
    `,
    [id],
  );
  return rows[0];
}

export async function upsertDailyReport(report: {
  id: string;
  userId: string;
  reportDate: string;
  status: DailyReport['status'];
  content?: string | null;
  reviewReason?: string | null;
  submittedAt?: string | null;
  reviewedAt?: string | null;
  reviewedBy?: string | null;
}): Promise<DailyReport> {
  const { rows } = await pool.query<DailyReport>(
    `
      INSERT INTO daily_reports (
        id,
        user_id,
        report_date,
        status,
        content,
        review_reason,
        created_at,
        updated_at,
        submitted_at,
        reviewed_at,
        reviewed_by
      )
      VALUES ($1, $2, $3, $4, $5, $6, NOW(), NOW(), $7, $8, $9)
      ON CONFLICT (user_id, report_date)
      DO UPDATE SET
        status = EXCLUDED.status,
        content = EXCLUDED.content,
        review_reason = EXCLUDED.review_reason,
        updated_at = NOW(),
        submitted_at = EXCLUDED.submitted_at,
        reviewed_at = EXCLUDED.reviewed_at,
        reviewed_by = EXCLUDED.reviewed_by
      RETURNING
        id,
        user_id AS "userId",
        report_date::text AS "reportDate",
        status,
        content,
        review_reason AS "reviewReason",
        created_at AS "createdAt",
        updated_at AS "updatedAt",
        submitted_at AS "submittedAt",
        reviewed_at AS "reviewedAt",
        reviewed_by AS "reviewedBy"
    `,
    [
      report.id,
      report.userId,
      report.reportDate,
      report.status,
      report.content ?? null,
      report.reviewReason ?? null,
      report.submittedAt ?? null,
      report.reviewedAt ?? null,
      report.reviewedBy ?? null,
    ],
  );
  return rows[0];
}

export async function updateDailyReportStatus(params: {
  id: string;
  status: DailyReport['status'];
  reviewedAt?: string | null;
  reviewedBy?: string | null;
  reviewReason?: string | null;
}): Promise<DailyReport | undefined> {
  const { rows } = await pool.query<DailyReport>(
    `
      UPDATE daily_reports
      SET status = $2,
          reviewed_at = $3,
          reviewed_by = $4,
          review_reason = $5,
          updated_at = NOW()
      WHERE id = $1
      RETURNING
        id,
        user_id AS "userId",
        report_date::text AS "reportDate",
        status,
        content,
        review_reason AS "reviewReason",
        created_at AS "createdAt",
        updated_at AS "updatedAt",
        submitted_at AS "submittedAt",
        reviewed_at AS "reviewedAt",
        reviewed_by AS "reviewedBy"
    `,
    [
      params.id,
      params.status,
      params.reviewedAt ?? null,
      params.reviewedBy ?? null,
      params.reviewReason ?? null,
    ],
  );
  return rows[0];
}

export async function updateProfileRecord(profile: {
  id: string;
  displayName: string;
  baseInfo: Record<string, unknown>;
  baseResume?: Record<string, unknown>;
  baseAdditionalBullets?: Record<string, number>;
  resumeTemplateId?: string | null;
}) {
  await pool.query(
    `
      UPDATE profiles
      SET display_name = $2,
          base_info = $3,
          base_resume = $4,
          base_additional_bullets = $5,
          resume_template_id = $6,
          updated_at = NOW()
      WHERE id = $1
    `,
    [
      profile.id,
      profile.displayName,
      JSON.stringify(profile.baseInfo ?? {}),
      JSON.stringify(profile.baseResume ?? {}),
      JSON.stringify(profile.baseAdditionalBullets ?? {}),
      profile.resumeTemplateId ?? null,
    ],
  );
}

export async function deleteProfile(id: string): Promise<boolean> {
  const { rowCount } = await pool.query('DELETE FROM profiles WHERE id = $1', [id]);
  return (rowCount ?? 0) > 0;
}

export async function listResumeTemplates(): Promise<ResumeTemplate[]> {
  const { rows } = await pool.query<ResumeTemplate>(
    `
      SELECT
        rt.id,
        rt.name,
        rt.description,
        rt.html,
        rt.created_by AS "createdBy",
        u.user_name AS "createdByName",
        rt.created_at AS "createdAt",
        rt.updated_at AS "updatedAt",
        COUNT(p.id)::int AS "profileCount"
      FROM resume_templates rt
      LEFT JOIN users u ON u.id = rt.created_by
      LEFT JOIN profiles p ON p.resume_template_id = rt.id
      GROUP BY rt.id, rt.name, rt.description, rt.html, rt.created_by, rt.created_at, rt.updated_at, u.user_name
      ORDER BY rt.updated_at DESC, rt.created_at DESC
    `,
  );
  return rows;
}

export async function findResumeTemplateById(id: string): Promise<ResumeTemplate | undefined> {
  const { rows } = await pool.query<ResumeTemplate>(
    `
      SELECT
        rt.id,
        rt.name,
        rt.description,
        rt.html,
        rt.created_by AS "createdBy",
        u.user_name AS "createdByName",
        rt.created_at AS "createdAt",
        rt.updated_at AS "updatedAt",
        COUNT(p.id)::int AS "profileCount"
      FROM resume_templates rt
      LEFT JOIN users u ON u.id = rt.created_by
      LEFT JOIN profiles p ON p.resume_template_id = rt.id
      WHERE rt.id = $1
      GROUP BY rt.id, rt.name, rt.description, rt.html, rt.created_by, rt.created_at, rt.updated_at, u.user_name
      LIMIT 1
    `,
    [id],
  );
  return rows[0];
}

export async function insertResumeTemplate(template: ResumeTemplate): Promise<ResumeTemplate> {
  const createdAt = template.createdAt ?? new Date().toISOString();
  const updatedAt = template.updatedAt ?? createdAt;
  const { rows } = await pool.query<ResumeTemplate>(
    `
      INSERT INTO resume_templates (
        id,
        name,
        description,
        html,
        created_by,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING
        id,
        name,
        description,
        html,
        created_by AS "createdBy",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
    `,
    [
      template.id,
      template.name,
      template.description ?? null,
      template.html,
      template.createdBy ?? null,
      createdAt,
      updatedAt,
    ],
  );
  const created = await findResumeTemplateById(template.id);
  if (created) return created;
  return { ...rows[0], profileCount: 0 };
}

export async function updateResumeTemplate(template: {
  id: string;
  name: string;
  description?: string | null;
  html: string;
}): Promise<ResumeTemplate | undefined> {
  const { rows } = await pool.query<ResumeTemplate>(
    `
      UPDATE resume_templates
      SET name = $2,
          description = $3,
          html = $4,
          updated_at = NOW()
      WHERE id = $1
      RETURNING
        id,
        name,
        description,
        html,
        created_by AS "createdBy",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
    `,
    [template.id, template.name, template.description ?? null, template.html],
  );
  if (!rows[0]) return undefined;
  const updated = await findResumeTemplateById(template.id);
  return updated ?? rows[0];
}

export async function deleteResumeTemplate(id: string): Promise<boolean> {
  const { rowCount } = await pool.query(
    'DELETE FROM resume_templates WHERE id = $1',
    [id],
  );
  return (rowCount ?? 0) > 0;
}

type TaskRow = Omit<Task, "tags" | "assignees"> & {
  tags?: string[] | null;
  assignees?: TaskAssignee[] | null;
};

function normalizeTaskRow(row: TaskRow): Task {
  return {
    ...row,
    tags: Array.isArray(row.tags) ? row.tags : [],
    assignees: Array.isArray(row.assignees) ? row.assignees : [],
  };
}

async function listTasksByQuery(
  clause: string,
  params: Array<string | number> = [],
): Promise<Task[]> {
  const { rows } = await pool.query<TaskRow>(
    `
      SELECT
        t.id,
        t.title,
        t.detail,
        t.status,
        t.priority,
        t.due_date::text AS "dueDate",
        t.project,
        t.notes,
        t.tags,
        t.created_by AS "createdBy",
        t.rejected_by AS "rejectedBy",
        t.rejected_at AS "rejectedAt",
        t.rejection_reason AS "rejectionReason",
        t.created_at AS "createdAt",
        t.updated_at AS "updatedAt",
        COALESCE(
          json_agg(
            json_build_object(
              'id', u.id,
              'userName', u.user_name,
              'email', u.email,
              'avatarUrl', u.avatar_url
            )
          ) FILTER (WHERE u.id IS NOT NULL),
          '[]'::json
        ) AS assignees
      FROM tasks t
      LEFT JOIN task_assignees ta ON ta.task_id = t.id
      LEFT JOIN users u ON u.id = ta.user_id
      ${clause}
      GROUP BY t.id
      ORDER BY t.due_date ASC, t.created_at DESC
    `,
    params,
  );
  return rows.map(normalizeTaskRow);
}

export async function listTasks(): Promise<Task[]> {
  return listTasksByQuery('');
}

export async function listTasksForUser(userId: string): Promise<Task[]> {
  return listTasksByQuery(
    `WHERE t.id IN (SELECT task_id FROM task_assignees WHERE user_id = $1)`,
    [userId],
  );
}

export async function findTaskById(id: string): Promise<Task | undefined> {
  const rows = await listTasksByQuery('WHERE t.id = $1', [id]);
  return rows[0];
}

export async function setTaskAssignees(taskId: string, userIds: string[]) {
  const cleaned = Array.from(
    new Set(userIds.map((id) => id.trim()).filter(Boolean)),
  );
  await pool.query('DELETE FROM task_assignees WHERE task_id = $1', [taskId]);
  if (cleaned.length === 0) return;
  const values = cleaned
    .map((_, index) => `($${index * 3 + 1}, $${index * 3 + 2}, $${index * 3 + 3})`)
    .join(', ');
  const params = cleaned.flatMap((userId) => [randomUUID(), taskId, userId]);
  await pool.query(
    `INSERT INTO task_assignees (id, task_id, user_id) VALUES ${values}`,
    params,
  );
}

export async function insertTask(task: {
  id: string;
  title: string;
  detail?: string | null;
  status: Task["status"];
  priority: Task["priority"];
  dueDate: string | null;
  project?: string | null;
  notes?: string | null;
  tags?: string[];
  createdBy?: string | null;
  rejectedBy?: string | null;
  rejectedAt?: string | null;
  rejectionReason?: string | null;
  assigneeIds?: string[];
}): Promise<Task> {
  const createdAt = new Date().toISOString();
  const updatedAt = createdAt;
  await pool.query(
    `
      INSERT INTO tasks (
        id,
        title,
        detail,
        status,
        priority,
        due_date,
        project,
        notes,
        tags,
        created_by,
        rejected_by,
        rejected_at,
        rejection_reason,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)
    `,
    [
      task.id,
      task.title,
      task.detail ?? null,
      task.status,
      task.priority,
      task.dueDate,
      task.project ?? null,
      task.notes ?? null,
      task.tags ?? [],
      task.createdBy ?? null,
      task.rejectedBy ?? null,
      task.rejectedAt ?? null,
      task.rejectionReason ?? null,
      createdAt,
      updatedAt,
    ],
  );
  await setTaskAssignees(task.id, task.assigneeIds ?? []);
  const created = await findTaskById(task.id);
  if (!created) {
    throw new Error('Task insert failed');
  }
  return created;
}

export async function updateTask(task: {
  id: string;
  title: string;
  detail?: string | null;
  status: Task["status"];
  priority: Task["priority"];
  dueDate: string | null;
  project?: string | null;
  notes?: string | null;
  tags?: string[];
  assigneeIds?: string[];
}): Promise<Task | undefined> {
  const { rowCount } = await pool.query(
    `
      UPDATE tasks
      SET title = $2,
          detail = $3,
          status = $4,
          priority = $5,
          due_date = $6,
          project = $7,
          notes = $8,
          tags = $9,
          updated_at = NOW()
      WHERE id = $1
    `,
    [
      task.id,
      task.title,
      task.detail ?? null,
      task.status,
      task.priority,
      task.dueDate,
      task.project ?? null,
      task.notes ?? null,
      task.tags ?? [],
    ],
  );
  if (!rowCount) return undefined;
  await setTaskAssignees(task.id, task.assigneeIds ?? []);
  return findTaskById(task.id);
}

export async function updateTaskNotes(
  taskId: string,
  notes: string | null,
): Promise<Task | undefined> {
  const { rowCount } = await pool.query(
    `
      UPDATE tasks
      SET notes = $2,
          updated_at = NOW()
      WHERE id = $1
    `,
    [taskId, notes ?? null],
  );
  if (!rowCount) return undefined;
  return findTaskById(taskId);
}

export async function deleteTask(id: string): Promise<boolean> {
  const { rowCount } = await pool.query('DELETE FROM tasks WHERE id = $1', [id]);
  return (rowCount ?? 0) > 0;
}

export async function rejectTask(
  taskId: string,
  reviewerId: string,
  reason?: string | null,
) {
  await pool.query(
    `
      UPDATE tasks
      SET rejected_by = $2,
          rejected_at = NOW(),
          rejection_reason = $3,
          updated_at = NOW()
      WHERE id = $1
    `,
    [taskId, reviewerId, reason ?? null],
  );
  await pool.query(
    `
      UPDATE task_assignment_requests
      SET status = 'rejected',
          reviewed_by = $2,
          reviewed_at = NOW(),
          rejection_reason = COALESCE(rejection_reason, $3),
          updated_at = NOW()
      WHERE task_id = $1 AND status = 'pending'
    `,
    [taskId, reviewerId, reason ?? 'Task rejected'],
  );
  return findTaskById(taskId);
}

export async function addTaskAssignee(taskId: string, userId: string) {
  await pool.query(
    `
      INSERT INTO task_assignees (id, task_id, user_id)
      VALUES ($1, $2, $3)
      ON CONFLICT (task_id, user_id) DO NOTHING
    `,
    [randomUUID(), taskId, userId],
  );
}

export async function updateTaskStatus(taskId: string, status: Task["status"]) {
  await pool.query(
    `
      UPDATE tasks
      SET status = $2,
          updated_at = NOW()
      WHERE id = $1
    `,
    [taskId, status],
  );
}

export async function upsertTaskAssignmentRequests(
  taskId: string,
  userIds: string[],
  requestedBy: string,
) {
  const cleaned = Array.from(
    new Set(userIds.map((id) => id.trim()).filter(Boolean)),
  );
  if (!cleaned.length) return;
  const values = cleaned
    .map(
      (_, index) =>
        `($${index * 4 + 1}, $${index * 4 + 2}, $${index * 4 + 3}, $${index * 4 + 4}, 'pending')`,
    )
    .join(', ');
  const params = cleaned.flatMap((userId) => [
    randomUUID(),
    taskId,
    userId,
    requestedBy,
  ]);
  await pool.query(
    `
      INSERT INTO task_assignment_requests (id, task_id, user_id, requested_by, status)
      VALUES ${values}
      ON CONFLICT (task_id, user_id)
      DO UPDATE SET
        status = 'pending',
        requested_by = EXCLUDED.requested_by,
        reviewed_by = NULL,
        reviewed_at = NULL,
        rejection_reason = NULL,
        updated_at = NOW()
    `,
    params,
  );
}

export async function listTaskAssignmentRequests(
  status?: TaskAssignmentRequestStatus,
): Promise<TaskAssignmentRequestDetail[]> {
  const params: Array<string | number> = [];
  let where = '';
  if (status) {
    params.push(status);
    where = 'WHERE tar.status = $1';
  }
  const { rows } = await pool.query<TaskAssignmentRequestDetail>(
    `
      SELECT
        tar.id,
        tar.task_id AS "taskId",
        tar.user_id AS "userId",
        tar.requested_by AS "requestedBy",
        tar.status,
        tar.reviewed_by AS "reviewedBy",
        tar.reviewed_at AS "reviewedAt",
        tar.rejection_reason AS "rejectionReason",
        tar.created_at AS "createdAt",
        tar.updated_at AS "updatedAt",
        t.title AS "taskTitle",
        u.user_name AS "assigneeName",
        u.email AS "assigneeEmail",
        r.user_name AS "requesterName",
        r.email AS "requesterEmail"
      FROM task_assignment_requests tar
      JOIN tasks t ON t.id = tar.task_id
      LEFT JOIN users u ON u.id = tar.user_id
      LEFT JOIN users r ON r.id = tar.requested_by
      ${where}
      ORDER BY tar.created_at DESC
    `,
    params,
  );
  return rows;
}

type TaskDoneRequestRow = TaskDoneRequestDetail;

async function listTaskDoneRequestsByQuery(
  clause: string,
  params: Array<string | number> = [],
): Promise<TaskDoneRequestDetail[]> {
  const { rows } = await pool.query<TaskDoneRequestRow>(
    `
      SELECT
        r.id,
        r.task_id AS "taskId",
        r.requested_by AS "requestedBy",
        r.status,
        r.reviewed_by AS "reviewedBy",
        r.reviewed_at AS "reviewedAt",
        r.rejection_reason AS "rejectionReason",
        r.created_at AS "createdAt",
        r.updated_at AS "updatedAt",
        t.title AS "taskTitle",
        requester.user_name AS "requesterName",
        requester.email AS "requesterEmail"
      FROM task_done_requests r
      JOIN tasks t ON t.id = r.task_id
      LEFT JOIN users requester ON requester.id = r.requested_by
      ${clause}
      ORDER BY r.created_at DESC
    `,
    params,
  );
  return rows;
}

export async function listTaskDoneRequests(
  status?: TaskDoneRequestStatus,
): Promise<TaskDoneRequestDetail[]> {
  if (!status) {
    return listTaskDoneRequestsByQuery('');
  }
  return listTaskDoneRequestsByQuery('WHERE r.status = $1', [status]);
}

export async function findTaskDoneRequestById(
  id: string,
): Promise<TaskDoneRequestDetail | undefined> {
  const rows = await listTaskDoneRequestsByQuery('WHERE r.id = $1', [id]);
  return rows[0];
}

export async function insertTaskDoneRequest(request: {
  id: string;
  taskId: string;
  requestedBy: string;
}): Promise<TaskDoneRequestDetail> {
  await pool.query(
    `
      INSERT INTO task_done_requests (
        id,
        task_id,
        requested_by,
        status,
        created_at,
        updated_at
      )
      VALUES ($1, $2, $3, $4, NOW(), NOW())
    `,
    [request.id, request.taskId, request.requestedBy, 'pending'],
  );
  const created = await findTaskDoneRequestById(request.id);
  if (!created) {
    throw new Error('Task done request insert failed');
  }
  return created;
}

export async function approveTaskDoneRequest(
  requestId: string,
  reviewerId: string,
): Promise<TaskDoneRequestDetail | undefined> {
  const existing = await findTaskDoneRequestById(requestId);
  if (!existing) return undefined;
  await pool.query(
    `
      UPDATE task_done_requests
      SET status = 'approved',
          reviewed_by = $2,
          reviewed_at = NOW(),
          rejection_reason = NULL,
          updated_at = NOW()
      WHERE id = $1
    `,
    [requestId, reviewerId],
  );
  // Get reviewer name and current task notes
  const { rows: userRows } = await pool.query<{ user_name: string }>(
    `SELECT user_name FROM users WHERE id = $1`,
    [reviewerId],
  );
  const reviewerName = userRows[0]?.user_name || 'Unknown';
  const { rows: taskRows } = await pool.query<{ notes: string | null }>(
    `SELECT notes FROM tasks WHERE id = $1`,
    [existing.taskId],
  );
  const currentNotes = taskRows[0]?.notes?.trim() || '';
  const stamp = new Date().toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
  const entry = `- ${stamp} - ${reviewerName}: Done request approved.`;
  const updatedNotes = currentNotes ? `${currentNotes}\n${entry}` : entry;
  await pool.query(
    `
      UPDATE tasks
      SET status = 'done',
          notes = $2,
          updated_at = NOW()
      WHERE id = $1
    `,
    [existing.taskId, updatedNotes],
  );
  return findTaskDoneRequestById(requestId);
}

export async function rejectTaskDoneRequest(
  requestId: string,
  reviewerId: string,
  reason?: string | null,
): Promise<TaskDoneRequestDetail | undefined> {
  const existing = await findTaskDoneRequestById(requestId);
  if (!existing) return undefined;
  await pool.query(
    `
      UPDATE task_done_requests
      SET status = 'rejected',
          reviewed_by = $2,
          reviewed_at = NOW(),
          rejection_reason = $3,
          updated_at = NOW()
      WHERE id = $1
    `,
    [requestId, reviewerId, reason ?? null],
  );
  // Get reviewer name and current task notes
  const { rows: userRows } = await pool.query<{ user_name: string }>(
    `SELECT user_name FROM users WHERE id = $1`,
    [reviewerId],
  );
  const reviewerName = userRows[0]?.user_name || 'Unknown';
  const { rows: taskRows } = await pool.query<{ notes: string | null }>(
    `SELECT notes FROM tasks WHERE id = $1`,
    [existing.taskId],
  );
  const currentNotes = taskRows[0]?.notes?.trim() || '';
  const stamp = new Date().toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
  const detail = reason
    ? `Done request rejected. Reason: ${reason}`
    : `Done request rejected.`;
  const entry = `- ${stamp} - ${reviewerName}: ${detail}`;
  const updatedNotes = currentNotes ? `${currentNotes}\n${entry}` : entry;
  await pool.query(
    `
      UPDATE tasks
      SET status = 'in_progress',
          notes = $2,
          updated_at = NOW()
      WHERE id = $1 AND status = 'in_review'
    `,
    [existing.taskId, updatedNotes],
  );
  return findTaskDoneRequestById(requestId);
}

export async function approveTaskAssignmentRequest(
  requestId: string,
  reviewerId: string,
) {
  const { rows } = await pool.query<{
    taskId: string;
    userId: string;
    requestedBy: string;
  }>(
    `
      UPDATE task_assignment_requests
      SET status = 'approved',
          reviewed_by = $2,
          reviewed_at = NOW(),
          rejection_reason = NULL,
          updated_at = NOW()
      WHERE id = $1
      RETURNING task_id AS "taskId", user_id AS "userId", requested_by AS "requestedBy"
    `,
    [requestId, reviewerId],
  );
  const request = rows[0];
  if (!request) return undefined;
  
  // Check if user is already assigned (to avoid duplicate notes)
  const { rows: existingAssignees } = await pool.query<{ user_id: string }>(
    `SELECT user_id FROM task_assignees WHERE task_id = $1 AND user_id = $2`,
    [request.taskId, request.userId],
  );
  const isNewAssignee = existingAssignees.length === 0;
  
  await addTaskAssignee(request.taskId, request.userId);
  
  // Add note if this is a new assignee
  if (isNewAssignee) {
    // Get reviewer name and new user name
    const { rows: reviewerRows } = await pool.query<{ user_name: string }>(
      `SELECT user_name FROM users WHERE id = $1`,
      [reviewerId],
    );
    const { rows: userRows } = await pool.query<{ user_name: string }>(
      `SELECT user_name FROM users WHERE id = $1`,
      [request.userId],
    );
    const reviewerName = reviewerRows[0]?.user_name || 'Unknown';
    const userName = userRows[0]?.user_name || 'Unknown';
    
    // Get current task notes
    const { rows: taskRows } = await pool.query<{ notes: string | null }>(
      `SELECT notes FROM tasks WHERE id = $1`,
      [request.taskId],
    );
    const currentNotes = taskRows[0]?.notes?.trim() || '';
    const stamp = new Date().toLocaleString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit",
    });
    const entry = `- ${stamp} - ${reviewerName}: New User join this task`;
    const updatedNotes = currentNotes ? `${currentNotes}\n${entry}` : entry;
    
    await pool.query(
      `
        UPDATE tasks
        SET notes = $2,
            updated_at = NOW()
        WHERE id = $1
      `,
      [request.taskId, updatedNotes],
    );
  }
  
  await pool.query(
    `
      UPDATE tasks
      SET status = 'in_progress',
          updated_at = NOW()
      WHERE id = $1 AND status = 'todo'
    `,
    [request.taskId],
  );
  return request;
}

export async function rejectTaskAssignmentRequest(
  requestId: string,
  reviewerId: string,
  reason?: string | null,
) {
  const { rows } = await pool.query<{
    taskId: string;
    userId: string;
    requestedBy: string;
  }>(
    `
      UPDATE task_assignment_requests
      SET status = 'rejected',
          reviewed_by = $2,
          reviewed_at = NOW(),
          rejection_reason = $3,
          updated_at = NOW()
      WHERE id = $1
      RETURNING task_id AS "taskId", user_id AS "userId", requested_by AS "requestedBy"
    `,
    [requestId, reviewerId, reason ?? null],
  );
  return rows[0];
}


export async function listAssignments(): Promise<Assignment[]> {
  const { rows } = await pool.query<Assignment>(
    `
      SELECT
        id AS "id",
        id AS "profileId",
        assigned_bidder_id AS "bidderUserId",
        assigned_by AS "assignedBy",
        assigned_at AS "assignedAt",
        NULL::TIMESTAMP AS "unassignedAt"
      FROM profiles
      WHERE assigned_bidder_id IS NOT NULL
      ORDER BY assigned_at DESC
    `,
  );
  return rows;
}

export async function findActiveAssignmentByProfile(
  profileId: string,
): Promise<Assignment | undefined> {
  const { rows } = await pool.query<Assignment>(
    `
      SELECT
        id AS "id",
        id AS "profileId",
        assigned_bidder_id AS "bidderUserId",
        assigned_by AS "assignedBy",
        assigned_at AS "assignedAt",
        NULL::TIMESTAMP AS "unassignedAt"
      FROM profiles
      WHERE id = $1 AND assigned_bidder_id IS NOT NULL
      LIMIT 1
    `,
    [profileId],
  );
  return rows[0];
}

export async function insertAssignmentRecord(assignment: Assignment) {
  await pool.query(
    `
      UPDATE profiles
      SET assigned_bidder_id = $2,
          assigned_by = $3,
          assigned_at = $4,
          updated_at = NOW()
      WHERE id = $1
    `,
    [assignment.profileId, assignment.bidderUserId, assignment.assignedBy, assignment.assignedAt],
  );
}

export async function closeAssignmentById(id: string): Promise<Assignment | undefined> {
  const { rows } = await pool.query<Assignment>(
    `
      UPDATE profiles
      SET assigned_bidder_id = NULL,
          assigned_by = NULL,
          assigned_at = NULL,
          updated_at = NOW()
      WHERE id = $1
      RETURNING
        id AS "id",
        id AS "profileId",
        NULL::UUID AS "bidderUserId",
        NULL::UUID AS "assignedBy",
        NULL::TIMESTAMP AS "assignedAt",
        NULL::TIMESTAMP AS "unassignedAt"
    `,
    [id],
  );
  return rows[0];
}

export type BidderSummary = {
  id: string;
  userName: string;
  email: string;
  profiles: { id: string; displayName: string }[];
};

export async function listBidderSummaries(): Promise<BidderSummary[]> {
  const { rows } = await pool.query<BidderSummary & { profiles?: { id: string; displayName: string }[] }>(
    `
      SELECT
        u.id,
        u.user_name AS "userName",
        u.email,
        COALESCE(
          json_agg(
            json_build_object('id', p.id, 'displayName', p.display_name)
          ) FILTER (WHERE p.id IS NOT NULL),
          '[]'::json
        ) AS profiles
      FROM users u
      LEFT JOIN profiles p ON p.assigned_bidder_id = u.id
      WHERE u.role = 'BIDDER' AND u.is_active IS NOT FALSE
      GROUP BY u.id, u.user_name, u.email
      ORDER BY u.user_name ASC
    `,
  );
  return rows.map((r) => ({
    id: r.id,
    userName: (r as any).userName,
    email: r.email,
    profiles: r.profiles ?? [],
  }));
}

const APPLICATION_SUMMARY_SELECT = `
  SELECT
    a.id,
    a.session_id AS "sessionId",
    a.bidder_user_id AS "bidderUserId",
    u.user_name AS "bidderName",
    u.email AS "bidderEmail",
    a.profile_id AS "profileId",
    p.display_name AS "profileDisplayName",
    a.resume_id AS "resumeId",
    NULL::text AS "resumeLabel",
    a.url AS "url",
    a.domain AS "domain",
    a.company AS "company",
    a.status AS "status",
    a.is_reviewed AS "isReviewed",
    a.reviewed_at AS "reviewedAt",
    a.reviewed_by AS "reviewedBy",
    reviewer.user_name AS "reviewerName",
    reviewer.email AS "reviewerEmail",
    a.created_at AS "createdAt"
  FROM applications a
  LEFT JOIN users u ON u.id = a.bidder_user_id
  LEFT JOIN profiles p ON p.id = a.profile_id
  LEFT JOIN users reviewer ON reviewer.id = a.reviewed_by
`;

export async function listApplications(): Promise<ApplicationSummary[]> {
  const { rows } = await pool.query<ApplicationSummary>(
    `
      ${APPLICATION_SUMMARY_SELECT}
      ORDER BY a.created_at DESC
    `,
  );
  return rows;
}

export async function updateApplicationReview(params: {
  id: string;
  isReviewed: boolean;
  reviewerId: string;
}): Promise<ApplicationSummary | undefined> {
  const reviewedAt = params.isReviewed ? new Date().toISOString() : null;
  const reviewedBy = params.isReviewed ? params.reviewerId : null;
  await pool.query(
    `
      UPDATE applications
      SET is_reviewed = $2,
          reviewed_by = $3,
          reviewed_at = $4
      WHERE id = $1
    `,
    [params.id, params.isReviewed, reviewedBy, reviewedAt],
  );
  const { rows } = await pool.query<ApplicationSummary>(
    `
      ${APPLICATION_SUMMARY_SELECT}
      WHERE a.id = $1
      LIMIT 1
    `,
    [params.id],
  );
  return rows[0];
}

export async function listApplicationsForBidder(
  bidderUserId: string,
): Promise<ApplicationSummary[]> {
  const { rows } = await pool.query<ApplicationSummary>(
    `
      ${APPLICATION_SUMMARY_SELECT}
      WHERE a.bidder_user_id = $1
      ORDER BY a.created_at DESC
    `,
    [bidderUserId],
  );
  return rows;
}

export async function findApplicationById(id: string): Promise<ApplicationSummary | undefined> {
  const { rows } = await pool.query<ApplicationSummary>(
    `
      ${APPLICATION_SUMMARY_SELECT}
      WHERE a.id = $1
    `,
    [id],
  );
  return rows[0];
}

export async function updateApplicationStatus(params: {
  id: string;
  status: "in_review" | "accepted" | "rejected";
}): Promise<ApplicationSummary | undefined> {
  await pool.query(
    `
      UPDATE applications
      SET status = $2
      WHERE id = $1
    `,
    [params.id, params.status],
  );
  const { rows } = await pool.query<ApplicationSummary>(
    `
      ${APPLICATION_SUMMARY_SELECT}
      WHERE a.id = $1
    `,
    [params.id],
  );
  return rows[0];
}

export async function listLabelAliases(): Promise<LabelAlias[]> {
  const { rows } = await pool.query<LabelAlias>(
    `
      SELECT
        id,
        canonical_key AS "canonicalKey",
        alias,
        normalized_alias AS "normalizedAlias",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
      FROM label_aliases
      ORDER BY created_at ASC
    `,
  );
  return rows;
}

export async function findLabelAliasById(id: string): Promise<LabelAlias | undefined> {
  const { rows } = await pool.query<LabelAlias>(
    `
      SELECT
        id,
        canonical_key AS "canonicalKey",
        alias,
        normalized_alias AS "normalizedAlias",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
      FROM label_aliases
      WHERE id = $1
      LIMIT 1
    `,
    [id],
  );
  return rows[0];
}

export async function findLabelAliasByNormalized(normalized: string): Promise<LabelAlias | undefined> {
  const { rows } = await pool.query<LabelAlias>(
    `
      SELECT
        id,
        canonical_key AS "canonicalKey",
        alias,
        normalized_alias AS "normalizedAlias",
        created_at AS "createdAt",
        updated_at AS "updatedAt"
      FROM label_aliases
      WHERE normalized_alias = $1
      LIMIT 1
    `,
    [normalized],
  );
  return rows[0];
}

export async function insertLabelAlias(alias: LabelAlias) {
  await pool.query(
    `
      INSERT INTO label_aliases (id, canonical_key, alias, normalized_alias, created_at, updated_at)
      VALUES ($1, $2, $3, $4, COALESCE($5, NOW()), COALESCE($6, NOW()))
      ON CONFLICT (normalized_alias) DO NOTHING
    `,
    [
      alias.id,
      alias.canonicalKey,
      alias.alias,
      alias.normalizedAlias,
      alias.createdAt ?? new Date().toISOString(),
      alias.updatedAt ?? new Date().toISOString(),
    ],
  );
}

export async function updateLabelAliasRecord(alias: LabelAlias) {
  await pool.query(
    `
      UPDATE label_aliases
      SET canonical_key = $2,
          alias = $3,
          normalized_alias = $4,
          updated_at = COALESCE($5, NOW())
      WHERE id = $1
    `,
    [alias.id, alias.canonicalKey, alias.alias, alias.normalizedAlias, alias.updatedAt ?? new Date().toISOString()],
  );
}

export async function deleteLabelAlias(id: string) {
  await pool.query('DELETE FROM label_aliases WHERE id = $1', [id]);
}

export async function listCommunityChannels(): Promise<CommunityThread[]> {
  const { rows } = await pool.query<CommunityThread>(
    `
      SELECT
        id,
        thread_type AS "threadType",
        name,
        description,
        created_by AS "createdBy",
        is_private AS "isPrivate",
        created_at AS "createdAt",
        last_message_at AS "lastMessageAt"
      FROM community_threads
      WHERE thread_type = 'CHANNEL'
      ORDER BY name ASC
    `,
  );
  return rows;
}

export async function updateCommunityChannel(channel: {
  id: string;
  name: string;
  nameKey: string;
  description: string | null;
}): Promise<CommunityThread | undefined> {
  const { rows } = await pool.query<CommunityThread>(
    `
      UPDATE community_threads
      SET name = $2,
          name_key = $3,
          description = $4
      WHERE id = $1 AND thread_type = 'CHANNEL'
      RETURNING
        id,
        thread_type AS "threadType",
        name,
        description,
        created_by AS "createdBy",
        is_private AS "isPrivate",
        created_at AS "createdAt",
        last_message_at AS "lastMessageAt"
    `,
    [channel.id, channel.name, channel.nameKey, channel.description],
  );
  return rows[0];
}

export async function deleteCommunityChannel(id: string): Promise<boolean> {
  const result = await pool.query(
    `
      DELETE FROM community_threads
      WHERE id = $1 AND thread_type = 'CHANNEL'
    `,
    [id],
  );
  return (result.rowCount ?? 0) > 0;
}

export async function listCommunityDmThreads(userId: string): Promise<CommunityThreadSummary[]> {
  const { rows } = await pool.query<CommunityThreadSummary & { participants?: CommunityThreadSummary['participants'] }>(
    `
      SELECT
        t.id,
        t.thread_type AS "threadType",
        t.name,
        t.description,
        t.is_private AS "isPrivate",
        t.created_at AS "createdAt",
        t.last_message_at AS "lastMessageAt",
        COALESCE(
          json_agg(
            json_build_object('id', u.id, 'userName', u.user_name, 'email', u.email, 'avatarUrl', u.avatar_url)
          ) FILTER (WHERE u.id IS NOT NULL),
          '[]'::json
        ) AS participants
      FROM community_threads t
      JOIN community_thread_members m ON m.thread_id = t.id AND m.user_id = $1
      LEFT JOIN community_thread_members m2 ON m2.thread_id = t.id
      LEFT JOIN users u ON u.id = m2.user_id AND u.id <> $1
      WHERE t.thread_type = 'DM'
      GROUP BY t.id
      ORDER BY COALESCE(t.last_message_at, t.created_at) DESC
    `,
    [userId],
  );
  return rows.map((row) => ({
    ...row,
    participants: row.participants ?? [],
  }));
}

export async function findCommunityChannelByKey(nameKey: string): Promise<CommunityThread | undefined> {
  const { rows } = await pool.query<CommunityThread>(
    `
      SELECT
        id,
        thread_type AS "threadType",
        name,
        description,
        created_by AS "createdBy",
        is_private AS "isPrivate",
        created_at AS "createdAt",
        last_message_at AS "lastMessageAt"
      FROM community_threads
      WHERE thread_type = 'CHANNEL' AND name_key = $1
      LIMIT 1
    `,
    [nameKey],
  );
  return rows[0];
}

export async function findCommunityThreadById(id: string): Promise<CommunityThread | undefined> {
  const { rows } = await pool.query<CommunityThread>(
    `
      SELECT
        id,
        thread_type AS "threadType",
        name,
        description,
        created_by AS "createdBy",
        is_private AS "isPrivate",
        created_at AS "createdAt",
        last_message_at AS "lastMessageAt"
      FROM community_threads
      WHERE id = $1
      LIMIT 1
    `,
    [id],
  );
  return rows[0];
}

export async function isCommunityThreadMember(threadId: string, userId: string): Promise<boolean> {
  const { rows } = await pool.query<{ ok: number }>(
    `
      SELECT 1 as ok
      FROM community_thread_members
      WHERE thread_id = $1 AND user_id = $2
      LIMIT 1
    `,
    [threadId, userId],
  );
  return rows.length > 0;
}

export async function insertCommunityThread(thread: {
  id: string;
  threadType: CommunityThread['threadType'];
  name?: string | null;
  nameKey?: string | null;
  description?: string | null;
  createdBy?: string | null;
  isPrivate?: boolean;
  createdAt?: string;
}): Promise<CommunityThread> {
  const createdAt = thread.createdAt ?? new Date().toISOString();
  const { rows } = await pool.query<CommunityThread>(
    `
      INSERT INTO community_threads (
        id,
        thread_type,
        name,
        name_key,
        description,
        created_by,
        is_private,
        created_at,
        last_message_at
      )
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, NULL)
      RETURNING
        id,
        thread_type AS "threadType",
        name,
        description,
        created_by AS "createdBy",
        is_private AS "isPrivate",
        created_at AS "createdAt",
        last_message_at AS "lastMessageAt"
    `,
    [
      thread.id,
      thread.threadType,
      thread.name ?? null,
      thread.nameKey ?? null,
      thread.description ?? null,
      thread.createdBy ?? null,
      thread.isPrivate ?? false,
      createdAt,
    ],
  );
  return rows[0];
}

export async function insertCommunityThreadMember(member: {
  id: string;
  threadId: string;
  userId: string;
  role?: string;
  joinedAt?: string;
}) {
  await pool.query(
    `
      INSERT INTO community_thread_members (id, thread_id, user_id, role, joined_at)
      VALUES ($1, $2, $3, $4, COALESCE($5, NOW()))
      ON CONFLICT (thread_id, user_id) DO NOTHING
    `,
    [member.id, member.threadId, member.userId, member.role ?? 'MEMBER', member.joinedAt ?? null],
  );
}

export async function findCommunityDmThreadId(userId: string, otherUserId: string): Promise<string | undefined> {
  const { rows } = await pool.query<{ id: string }>(
    `
      SELECT t.id
      FROM community_threads t
      JOIN community_thread_members m1 ON m1.thread_id = t.id AND m1.user_id = $1
      JOIN community_thread_members m2 ON m2.thread_id = t.id AND m2.user_id = $2
      WHERE t.thread_type = 'DM'
      LIMIT 1
    `,
    [userId, otherUserId],
  );
  return rows[0]?.id;
}

export async function getCommunityDmThreadSummary(
  threadId: string,
  userId: string,
): Promise<CommunityThreadSummary | undefined> {
  const { rows } = await pool.query<CommunityThreadSummary & { participants?: CommunityThreadSummary['participants'] }>(
    `
      SELECT
        t.id,
        t.thread_type AS "threadType",
        t.name,
        t.description,
        t.is_private AS "isPrivate",
        t.created_at AS "createdAt",
        t.last_message_at AS "lastMessageAt",
        COALESCE(
          json_agg(
            json_build_object('id', u.id, 'userName', u.user_name, 'email', u.email, 'avatarUrl', u.avatar_url)
          ) FILTER (WHERE u.id IS NOT NULL),
          '[]'::json
        ) AS participants
      FROM community_threads t
      JOIN community_thread_members m ON m.thread_id = t.id AND m.user_id = $1
      LEFT JOIN community_thread_members m2 ON m2.thread_id = t.id
      LEFT JOIN users u ON u.id = m2.user_id AND u.id <> $1
      WHERE t.thread_type = 'DM' AND t.id = $2
      GROUP BY t.id
      LIMIT 1
    `,
    [userId, threadId],
  );
  const row = rows[0];
  if (!row) return undefined;
  return { ...row, participants: row.participants ?? [] };
}

export async function listCommunityMessages(threadId: string): Promise<CommunityMessage[]> {
  const { rows } = await pool.query<CommunityMessage>(
    `
      SELECT
        m.id,
        m.thread_id AS "threadId",
        m.sender_id AS "senderId",
        u.user_name AS "senderName",
        u.avatar_url AS "senderAvatarUrl",
        m.body,
        m.created_at AS "createdAt"
      FROM community_messages m
      LEFT JOIN users u ON u.id = m.sender_id
      WHERE m.thread_id = $1
      ORDER BY m.created_at ASC
      LIMIT 200
    `,
    [threadId],
  );
  return rows;
}

export async function listCommunityThreadMemberIds(threadId: string): Promise<string[]> {
  const { rows } = await pool.query<{ userId: string }>(
    `
      SELECT user_id AS "userId"
      FROM community_thread_members
      WHERE thread_id = $1
    `,
    [threadId],
  );
  return rows.map((row) => row.userId);
}

export async function insertCommunityMessage(message: CommunityMessage): Promise<CommunityMessage> {
  const createdAt = message.createdAt ?? new Date().toISOString();
  const { rows } = await pool.query<CommunityMessage>(
    `
      WITH inserted AS (
        INSERT INTO community_messages (
          id, thread_id, sender_id, body, reply_to_message_id, 
          is_edited, is_deleted, created_at
        )
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
        RETURNING 
          id, thread_id, sender_id, body, reply_to_message_id,
          is_edited, edited_at, is_deleted, deleted_at, created_at
      )
      SELECT
        inserted.id,
        inserted.thread_id AS "threadId",
        inserted.sender_id AS "senderId",
        u.user_name AS "senderName",
        u.avatar_url AS "senderAvatarUrl",
        inserted.body,
        inserted.reply_to_message_id AS "replyToMessageId",
        inserted.is_edited AS "isEdited",
        inserted.edited_at AS "editedAt",
        inserted.is_deleted AS "isDeleted",
        inserted.deleted_at AS "deletedAt",
        inserted.created_at AS "createdAt"
      FROM inserted
      LEFT JOIN users u ON u.id = inserted.sender_id
    `,
    [
      message.id, 
      message.threadId, 
      message.senderId, 
      message.body, 
      message.replyToMessageId ?? null,
      message.isEdited ?? false,
      message.isDeleted ?? false,
      createdAt
    ],
  );
  await pool.query('UPDATE community_threads SET last_message_at = $2 WHERE id = $1', [
    message.threadId,
    createdAt,
  ]);
  return rows[0];
}

// Message Attachments
export async function insertMessageAttachment(attachment: MessageAttachment): Promise<MessageAttachment> {
  const { rows } = await pool.query<MessageAttachment>(
    `
      INSERT INTO community_message_attachments 
        (id, message_id, file_name, file_url, file_size, mime_type, thumbnail_url, width, height, created_at)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
      RETURNING 
        id, message_id AS "messageId", file_name AS "fileName", file_url AS "fileUrl", 
        file_size AS "fileSize", mime_type AS "mimeType", thumbnail_url AS "thumbnailUrl",
        width, height, created_at AS "createdAt"
    `,
    [
      attachment.id,
      attachment.messageId,
      attachment.fileName,
      attachment.fileUrl,
      attachment.fileSize,
      attachment.mimeType,
      attachment.thumbnailUrl ?? null,
      attachment.width ?? null,
      attachment.height ?? null,
      attachment.createdAt ?? new Date().toISOString(),
    ],
  );
  return rows[0];
}

export async function listMessageAttachments(messageId: string): Promise<MessageAttachment[]> {
  const { rows } = await pool.query<MessageAttachment>(
    `
      SELECT 
        id, message_id AS "messageId", file_name AS "fileName", file_url AS "fileUrl",
        file_size AS "fileSize", mime_type AS "mimeType", thumbnail_url AS "thumbnailUrl",
        width, height, created_at AS "createdAt"
      FROM community_message_attachments
      WHERE message_id = $1
      ORDER BY created_at ASC
    `,
    [messageId],
  );
  return rows;
}

// Message Reactions
export async function addMessageReaction(reaction: MessageReaction): Promise<MessageReaction> {
  const { rows } = await pool.query<MessageReaction>(
    `
      INSERT INTO community_message_reactions (id, message_id, user_id, emoji, created_at)
      VALUES ($1, $2, $3, $4, $5)
      ON CONFLICT (message_id, user_id, emoji) DO NOTHING
      RETURNING 
        id, message_id AS "messageId", user_id AS "userId", emoji, created_at AS "createdAt"
    `,
    [reaction.id, reaction.messageId, reaction.userId, reaction.emoji, reaction.createdAt ?? new Date().toISOString()],
  );
  return rows[0];
}

export async function removeMessageReaction(messageId: string, userId: string, emoji: string): Promise<boolean> {
  const { rowCount } = await pool.query(
    'DELETE FROM community_message_reactions WHERE message_id = $1 AND user_id = $2 AND emoji = $3',
    [messageId, userId, emoji],
  );
  return (rowCount ?? 0) > 0;
}

export async function listMessageReactions(messageId: string, currentUserId?: string): Promise<ReactionSummary[]> {
  const { rows } = await pool.query<{ emoji: string; count: number; userIds: string[] }>(
    `
      SELECT emoji, COUNT(*)::int AS count, ARRAY_AGG(user_id) AS "userIds"
      FROM community_message_reactions
      WHERE message_id = $1
      GROUP BY emoji
      ORDER BY count DESC, emoji
    `,
    [messageId],
  );
  return rows.map((row) => ({
    emoji: row.emoji,
    count: row.count,
    userIds: row.userIds,
    hasCurrentUser: currentUserId ? row.userIds.includes(currentUserId) : false,
  }));
}

// Unread Messages
export async function getUnreadInfo(threadId: string, userId: string): Promise<UnreadInfo | null> {
  const { rows } = await pool.query<UnreadInfo>(
    `
      SELECT 
        thread_id AS "threadId", 
        unread_count AS "unreadCount", 
        last_read_message_id AS "lastReadMessageId",
        last_read_at AS "lastReadAt"
      FROM community_unread_messages
      WHERE thread_id = $1 AND user_id = $2
    `,
    [threadId, userId],
  );
  return rows[0] ?? null;
}

export async function markThreadAsRead(threadId: string, userId: string, messageId?: string): Promise<void> {
  await pool.query(
    `
      INSERT INTO community_unread_messages (id, thread_id, user_id, last_read_message_id, unread_count, last_read_at, updated_at)
      VALUES ($1, $2, $3, $4, 0, NOW(), NOW())
      ON CONFLICT (thread_id, user_id) 
      DO UPDATE SET 
        last_read_message_id = COALESCE($4, community_unread_messages.last_read_message_id),
        unread_count = 0, 
        last_read_at = NOW(), 
        updated_at = NOW()
    `,
    [randomUUID(), threadId, userId, messageId ?? null],
  );
}

export async function incrementUnreadCount(threadId: string, excludeUserId: string): Promise<void> {
  await pool.query(
    `
      INSERT INTO community_unread_messages (id, thread_id, user_id, unread_count, updated_at)
      SELECT $1, $2, user_id, 1, NOW()
      FROM community_thread_members
      WHERE thread_id = $2 AND user_id != $3
      ON CONFLICT (thread_id, user_id)
      DO UPDATE SET unread_count = community_unread_messages.unread_count + 1, updated_at = NOW()
    `,
    [randomUUID(), threadId, excludeUserId],
  );
}

// Pagination support for messages
export async function listCommunityMessagesWithPagination(
  threadId: string,
  options: { limit?: number; before?: string; after?: string } = {},
): Promise<CommunityMessageExtended[]> {
  const limit = options.limit ?? 50;
  let query = `
    SELECT 
      m.id, m.thread_id AS "threadId", m.sender_id AS "senderId", 
      u.user_name AS "senderName", u.avatar_url AS "senderAvatarUrl", m.body, m.reply_to_message_id AS "replyToMessageId",
      m.is_edited AS "isEdited", m.edited_at AS "editedAt",
      m.is_deleted AS "isDeleted", m.deleted_at AS "deletedAt",
      m.created_at AS "createdAt"
    FROM community_messages m
    LEFT JOIN users u ON u.id = m.sender_id
    WHERE m.thread_id = $1
  `;
  const params: any[] = [threadId];
  
  if (options.before) {
    query += ` AND m.created_at < (SELECT created_at FROM community_messages WHERE id = $2)`;
    params.push(options.before);
  } else if (options.after) {
    query += ` AND m.created_at > (SELECT created_at FROM community_messages WHERE id = $2)`;
    params.push(options.after);
  }
  
  query += ` ORDER BY m.created_at DESC LIMIT $${params.length + 1}`;
  params.push(limit);
  
  const { rows } = await pool.query<CommunityMessageExtended>(query, params);
  return rows.reverse();
}

export async function getMessageById(messageId: string): Promise<CommunityMessage | null> {
  const { rows } = await pool.query<CommunityMessage>(
    `
      SELECT 
        m.id, m.thread_id AS "threadId", m.sender_id AS "senderId",
        u.user_name AS "senderName", u.avatar_url AS "senderAvatarUrl", m.body, m.reply_to_message_id AS "replyToMessageId",
        m.is_edited AS "isEdited", m.edited_at AS "editedAt",
        m.is_deleted AS "isDeleted", m.deleted_at AS "deletedAt",
        m.created_at AS "createdAt"
      FROM community_messages m
      LEFT JOIN users u ON u.id = m.sender_id
      WHERE m.id = $1
    `,
    [messageId],
  );
  return rows[0] ?? null;
}

// Edit and Delete messages
export async function editMessage(messageId: string, body: string): Promise<CommunityMessage | null> {
  const { rows } = await pool.query<CommunityMessage>(
    `
      UPDATE community_messages 
      SET body = $2, is_edited = TRUE, edited_at = NOW()
      WHERE id = $1
      RETURNING 
        id, thread_id AS "threadId", sender_id AS "senderId", body,
        reply_to_message_id AS "replyToMessageId", is_edited AS "isEdited", edited_at AS "editedAt",
        is_deleted AS "isDeleted", deleted_at AS "deletedAt", created_at AS "createdAt"
    `,
    [messageId, body],
  );
  return rows[0] ?? null;
}

export async function deleteMessage(messageId: string): Promise<boolean> {
  const { rowCount } = await pool.query(
    `
      UPDATE community_messages 
      SET is_deleted = TRUE, deleted_at = NOW(), body = '[deleted]'
      WHERE id = $1
    `,
    [messageId],
  );
  return (rowCount ?? 0) > 0;
}

// Pinned Messages
export async function pinMessage(threadId: string, messageId: string, userId: string): Promise<PinnedMessage | null> {
  const { rows } = await pool.query<PinnedMessage>(
    `
      WITH inserted AS (
        INSERT INTO community_pinned_messages (id, thread_id, message_id, pinned_by, pinned_at)
        VALUES ($1, $2, $3, $4, NOW())
        ON CONFLICT (thread_id, message_id) DO NOTHING
        RETURNING id, thread_id, message_id, pinned_by, pinned_at
      )
      SELECT 
        inserted.id, 
        inserted.thread_id AS "threadId", 
        inserted.message_id AS "messageId", 
        inserted.pinned_by AS "pinnedBy", 
        inserted.pinned_at AS "pinnedAt",
        jsonb_build_object(
          'id', m.id,
          'threadId', m.thread_id,
          'senderId', m.sender_id,
          'senderName', u.user_name,
          'senderAvatarUrl', u.avatar_url,
          'body', m.body,
          'createdAt', m.created_at,
          'isEdited', m.is_edited,
          'isDeleted', m.is_deleted
        ) as message
      FROM inserted
      JOIN community_messages m ON m.id = inserted.message_id
      JOIN users u ON u.id = m.sender_id
    `,
    [randomUUID(), threadId, messageId, userId],
  );
  return rows[0] || null;
}

export async function unpinMessage(threadId: string, messageId: string): Promise<boolean> {
  const { rowCount } = await pool.query(
    'DELETE FROM community_pinned_messages WHERE thread_id = $1 AND message_id = $2',
    [threadId, messageId],
  );
  return (rowCount ?? 0) > 0;
}

export async function listPinnedMessages(threadId: string): Promise<PinnedMessage[]> {
  const { rows } = await pool.query<PinnedMessage>(
    `
      SELECT 
        p.id, 
        p.thread_id AS "threadId", 
        p.message_id AS "messageId", 
        p.pinned_by AS "pinnedBy", 
        p.pinned_at AS "pinnedAt",
        jsonb_build_object(
          'id', m.id,
          'threadId', m.thread_id,
          'senderId', m.sender_id,
          'senderName', u.user_name,
          'senderAvatarUrl', u.avatar_url,
          'body', m.body,
          'createdAt', m.created_at,
          'isEdited', m.is_edited,
          'isDeleted', m.is_deleted
        ) as message
      FROM community_pinned_messages p
      JOIN community_messages m ON m.id = p.message_id
      JOIN users u ON u.id = m.sender_id
      WHERE p.thread_id = $1
      ORDER BY p.pinned_at DESC
    `,
    [threadId],
  );
  return rows;
}

// User Presence
export async function updateUserPresence(userId: string, status: 'online' | 'away' | 'busy' | 'offline'): Promise<void> {
  await pool.query(
    `
      INSERT INTO community_user_presence (user_id, status, last_seen_at, updated_at)
      VALUES ($1, $2, NOW(), NOW())
      ON CONFLICT (user_id) 
      DO UPDATE SET status = $2, last_seen_at = NOW(), updated_at = NOW()
    `,
    [userId, status],
  );
}

export async function getUserPresence(userId: string): Promise<UserPresence | null> {
  const { rows } = await pool.query<UserPresence>(
    `
      SELECT user_id AS "userId", status, last_seen_at AS "lastSeenAt"
      FROM community_user_presence
      WHERE user_id = $1
    `,
    [userId],
  );
  return rows[0] ?? null;
}

export async function listUserPresences(userIds: string[]): Promise<UserPresence[]> {
  if (userIds.length === 0) return [];
  const { rows } = await pool.query<UserPresence>(
    `
      SELECT user_id AS "userId", status, last_seen_at AS "lastSeenAt"
      FROM community_user_presence
      WHERE user_id = ANY($1)
    `,
    [userIds],
  );
  return rows;
}
// Message Read Receipts
export async function addMessageReadReceipt(messageId: string, userId: string): Promise<void> {
  await pool.query(
    `
      INSERT INTO community_message_read_receipts (id, message_id, user_id, read_at)
      VALUES ($1, $2, $3, NOW())
      ON CONFLICT (message_id, user_id) DO UPDATE SET read_at = NOW()
    `,
    [randomUUID(), messageId, userId],
  );
}

export async function getMessageReadReceipts(messageId: string): Promise<{ totalReaders: number; userIds: string[] }> {
  const { rows } = await pool.query<{ userId: string }>(
    `
      SELECT user_id AS "userId"
      FROM community_message_read_receipts
      WHERE message_id = $1
      ORDER BY read_at ASC
    `,
    [messageId],
  );
  return {
    totalReaders: rows.length,
    userIds: rows.map(r => r.userId),
  };
}

export async function bulkAddReadReceipts(messageIds: string[], userId: string): Promise<void> {
  if (messageIds.length === 0) return;
  
  const values = messageIds.map((_, i) => `($${i * 3 + 1}, $${i * 3 + 2}, $${i * 3 + 3}, NOW())`).join(', ');
  const params = messageIds.flatMap(msgId => [randomUUID(), msgId, userId]);
  
  await pool.query(
    `
      INSERT INTO community_message_read_receipts (id, message_id, user_id, read_at)
      VALUES ${values}
      ON CONFLICT (message_id, user_id) DO UPDATE SET read_at = NOW()
    `,
    params,
  );
}
