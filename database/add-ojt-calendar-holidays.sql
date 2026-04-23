-- Run in Supabase SQL editor if your project was created before calendar/holiday support.
-- Company-wide no-work days and per-trainee overrides for weekends/holidays.

CREATE TABLE IF NOT EXISTS ojt_holidays (
    date TEXT PRIMARY KEY,
    label TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE ojt_holidays ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_holidays" ON ojt_holidays;
CREATE POLICY "Allow all on ojt_holidays" ON ojt_holidays FOR ALL USING (true) WITH CHECK (true);

CREATE TABLE IF NOT EXISTS ojt_attendance_day_allow (
    trainee_id UUID NOT NULL REFERENCES ojt_trainees(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    PRIMARY KEY (trainee_id, date)
);

ALTER TABLE ojt_attendance_day_allow ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_attendance_day_allow" ON ojt_attendance_day_allow;
CREATE POLICY "Allow all on ojt_attendance_day_allow" ON ojt_attendance_day_allow FOR ALL USING (true) WITH CHECK (true);
