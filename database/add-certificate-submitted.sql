-- Run once in Supabase SQL Editor: flag when the admin has submitted the trainee's certificate.
ALTER TABLE ojt_trainees ADD COLUMN IF NOT EXISTS certificate_submitted_at TIMESTAMPTZ;

COMMENT ON COLUMN ojt_trainees.certificate_submitted_at IS 'Set when admin clicks Submit on Certificate; trainee dashboard shows certificate only after this is set.';
