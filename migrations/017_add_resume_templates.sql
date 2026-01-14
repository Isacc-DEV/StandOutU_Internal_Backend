CREATE TABLE IF NOT EXISTS resume_templates (
  id UUID PRIMARY KEY,
  name TEXT NOT NULL,
  description TEXT,
  html TEXT NOT NULL,
  created_by UUID REFERENCES users(id),
  created_at TIMESTAMP DEFAULT NOW(),
  updated_at TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_resume_templates_updated ON resume_templates(updated_at);
CREATE INDEX IF NOT EXISTS idx_resume_templates_name ON resume_templates(name);
