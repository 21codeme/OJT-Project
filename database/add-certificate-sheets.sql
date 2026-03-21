-- Run once in Supabase: multiple certificate/frame uploads per trainee (no longer replacing a single row).
CREATE TABLE IF NOT EXISTS ojt_trainee_certificate_sheets (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    trainee_id UUID NOT NULL REFERENCES ojt_trainees(id) ON DELETE CASCADE,
    file_name TEXT,
    file_data TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_cert_sheets_trainee ON ojt_trainee_certificate_sheets(trainee_id);
CREATE INDEX IF NOT EXISTS idx_cert_sheets_created ON ojt_trainee_certificate_sheets(created_at DESC);

ALTER TABLE ojt_trainee_certificate_sheets ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_trainee_certificate_sheets" ON ojt_trainee_certificate_sheets;
CREATE POLICY "Allow all on ojt_trainee_certificate_sheets" ON ojt_trainee_certificate_sheets FOR ALL USING (true) WITH CHECK (true);
