-- Supabase Database Setup for Class Schedule (Lab Schedule)
-- Run this SQL in your Supabase project: SQL Editor > New Query > Paste > Run

-- Sheets (COMPUTER LABORATORY, MULTIMEDIA AND SPEECH LABORATORY, etc.)
CREATE TABLE IF NOT EXISTS class_schedule_sheets (
    id BIGSERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- Schedule entries per sheet (day, time_slot, type, instructor, course, code)
CREATE TABLE IF NOT EXISTS class_schedule_entries (
    id BIGSERIAL PRIMARY KEY,
    sheet_id BIGINT NOT NULL REFERENCES class_schedule_sheets(id) ON DELETE CASCADE,
    day TEXT NOT NULL,
    time_slot TEXT NOT NULL,
    type TEXT,
    instructor TEXT,
    course TEXT,
    code TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_class_schedule_entries_sheet_id ON class_schedule_entries(sheet_id);
CREATE INDEX IF NOT EXISTS idx_class_schedule_entries_day ON class_schedule_entries(sheet_id, day);

-- Row Level Security
ALTER TABLE class_schedule_sheets ENABLE ROW LEVEL SECURITY;
ALTER TABLE class_schedule_entries ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow all on class_schedule_sheets" ON class_schedule_sheets
    FOR ALL USING (true) WITH CHECK (true);

CREATE POLICY "Allow all on class_schedule_entries" ON class_schedule_entries
    FOR ALL USING (true) WITH CHECK (true);

-- Auto-update updated_at
CREATE OR REPLACE FUNCTION update_class_schedule_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER class_schedule_sheets_updated_at
    BEFORE UPDATE ON class_schedule_sheets
    FOR EACH ROW EXECUTE FUNCTION update_class_schedule_updated_at();

CREATE TRIGGER class_schedule_entries_updated_at
    BEFORE UPDATE ON class_schedule_entries
    FOR EACH ROW EXECUTE FUNCTION update_class_schedule_updated_at();
