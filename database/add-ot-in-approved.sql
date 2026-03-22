-- OT time-in must be approved by admin before OT hours count toward totals.
-- Run once in Supabase SQL Editor after ojt-tables.sql.

ALTER TABLE ojt_attendance ADD COLUMN IF NOT EXISTS ot_in_approved BOOLEAN NOT NULL DEFAULT false;

-- Existing rows that already have OT In: treat as approved (no disruption).
UPDATE ojt_attendance
SET ot_in_approved = true
WHERE ot_in_time IS NOT NULL AND length(trim(ot_in_time)) > 0;

COMMENT ON COLUMN ojt_attendance.ot_in_approved IS 'When false, trainee OT In is pending; OT hours do not count until admin approves.';
