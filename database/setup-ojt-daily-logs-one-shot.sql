-- =============================================================================
-- OJT Daily Logs — fix 404 (REST) + 400 (Storage) in one run
-- =============================================================================
-- Run in Supabase: SQL Editor → New query → paste → Run.
--
-- The `ojt_trainees` table must already exist (from ojt-tables.sql).
-- If the full OJT schema is not there yet, run `ojt-tables.sql` first.
-- =============================================================================

-- 1) Table: ojt_daily_logs (PostgREST 404 if this is missing)
CREATE TABLE IF NOT EXISTS ojt_daily_logs (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    trainee_id UUID NOT NULL REFERENCES ojt_trainees(id) ON DELETE CASCADE,
    log_date TEXT NOT NULL,
    log_text TEXT NOT NULL,
    photo_url TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE ojt_daily_logs ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Allow all on ojt_daily_logs" ON ojt_daily_logs;
CREATE POLICY "Allow all on ojt_daily_logs" ON ojt_daily_logs
  FOR ALL USING (true) WITH CHECK (true);

-- 2) Storage bucket (400 if the bucket does not exist)
INSERT INTO storage.buckets (id, name, public)
VALUES ('ojt-daily-logs', 'ojt-daily-logs', true)
ON CONFLICT (id) DO UPDATE SET public = true;

-- 3) Storage RLS policies for anon key (same as ojt-daily-logs-storage.sql)
DROP POLICY IF EXISTS "Public insert on ojt-daily-logs" ON storage.objects;
CREATE POLICY "Public insert on ojt-daily-logs"
ON storage.objects FOR INSERT
TO public
WITH CHECK (bucket_id = 'ojt-daily-logs');

DROP POLICY IF EXISTS "Public read on ojt-daily-logs" ON storage.objects;
CREATE POLICY "Public read on ojt-daily-logs"
ON storage.objects FOR SELECT
TO public
USING (bucket_id = 'ojt-daily-logs');

DROP POLICY IF EXISTS "Public update on ojt-daily-logs" ON storage.objects;
CREATE POLICY "Public update on ojt-daily-logs"
ON storage.objects FOR UPDATE
TO public
USING (bucket_id = 'ojt-daily-logs')
WITH CHECK (bucket_id = 'ojt-daily-logs');

DROP POLICY IF EXISTS "Public delete on ojt-daily-logs" ON storage.objects;
CREATE POLICY "Public delete on ojt-daily-logs"
ON storage.objects FOR DELETE
TO public
USING (bucket_id = 'ojt-daily-logs');

-- 4) Refresh PostgREST schema cache (avoids "not in schema cache" after creating the table)
NOTIFY pgrst, 'reload schema';
