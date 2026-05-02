-- Run once (server also applies this on startup if tables are missing).
CREATE TABLE IF NOT EXISTS file_uploads (
  id BIGSERIAL PRIMARY KEY,
  original_name TEXT NOT NULL,
  content_type TEXT,
  byte_size BIGINT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  row_count INTEGER NOT NULL DEFAULT 0,
  sheet_names TEXT[] NOT NULL DEFAULT '{}',
  first_sheet_name TEXT,
  column_headers TEXT[] NOT NULL DEFAULT '{}'
);

CREATE TABLE IF NOT EXISTS file_upload_rows (
  id BIGSERIAL PRIMARY KEY,
  upload_id BIGINT NOT NULL REFERENCES file_uploads (id) ON DELETE CASCADE,
  row_index INTEGER NOT NULL,
  cells JSONB NOT NULL DEFAULT '{}'::jsonb,
  UNIQUE (upload_id, row_index)
);

CREATE INDEX IF NOT EXISTS idx_file_upload_rows_upload_id ON file_upload_rows (upload_id);

ALTER TABLE file_uploads ADD COLUMN IF NOT EXISTS column_headers TEXT[] NOT NULL DEFAULT '{}';

-- Invite links for contributors to edit spillover / bug fields for one upload
CREATE TABLE IF NOT EXISTS invites (
  id BIGSERIAL PRIMARY KEY,
  token UUID NOT NULL UNIQUE,
  upload_id BIGINT NOT NULL REFERENCES file_uploads (id) ON DELETE CASCADE,
  invitee_email TEXT,
  filter_by_assignee BOOLEAN NOT NULL DEFAULT false,
  assignee_scope TEXT,
  note TEXT,
  expires_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_invites_upload_id ON invites (upload_id);

ALTER TABLE invites ADD COLUMN IF NOT EXISTS assignee_scope TEXT;

-- Sprint threshold stored when the invite was created (same as Send invite); used to scope contributor rows.
ALTER TABLE invites ADD COLUMN IF NOT EXISTS sprint_threshold TEXT;

-- Shared enrichment keyed by upload + Jira issue key (filled by analysts or invitees)
CREATE TABLE IF NOT EXISTS issue_field_edits (
  upload_id BIGINT NOT NULL REFERENCES file_uploads (id) ON DELETE CASCADE,
  issue_key TEXT NOT NULL,
  spillover_reason TEXT NOT NULL DEFAULT '',
  spillover_category TEXT NOT NULL DEFAULT '',
  prod TEXT NOT NULL DEFAULT '',
  rca TEXT NOT NULL DEFAULT '',
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  PRIMARY KEY (upload_id, issue_key)
);

CREATE INDEX IF NOT EXISTS idx_issue_field_edits_upload ON issue_field_edits (upload_id);
