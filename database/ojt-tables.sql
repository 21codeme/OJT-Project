-- OJT Trainees, Documents, Attendance, Admin — run sa Supabase SQL Editor
-- Creates tables for trainee accounts, uploaded documents, attendance records, and admin.
-- Mga dokumento (file): gumawa ng public Storage bucket na ojt-trainee-documents at i-run ang ojt-documents-storage.sql.

-- 1) OJT Trainees (profile from create account)
CREATE TABLE IF NOT EXISTS ojt_trainees (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    email TEXT UNIQUE NOT NULL,
    username TEXT,
    full_name TEXT,
    student_id TEXT,
    course TEXT,
    year_level TEXT,
    school_name TEXT,
    department TEXT,
    coordinator TEXT,
    company_name TEXT,
    company_address TEXT,
    supervisor TEXT,
    position_role TEXT,
    contact_number TEXT,
    emergency_contact TEXT,
    training_hours TEXT,
    start_date TEXT,
    end_date TEXT,
    picture_url TEXT,
    password_hash TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Add password column if table already exists (run once)
ALTER TABLE ojt_trainees ADD COLUMN IF NOT EXISTS password_hash TEXT;

ALTER TABLE ojt_trainees ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_trainees" ON ojt_trainees;
CREATE POLICY "Allow all on ojt_trainees" ON ojt_trainees FOR ALL USING (true) WITH CHECK (true);

-- 2) OJT Trainee Documents (uploaded files per trainee)
CREATE TABLE IF NOT EXISTS ojt_trainee_documents (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    trainee_id UUID NOT NULL REFERENCES ojt_trainees(id) ON DELETE CASCADE,
    doc_type TEXT NOT NULL,
    file_name TEXT,
    file_data TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(trainee_id, doc_type)
);

ALTER TABLE ojt_trainee_documents ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_trainee_documents" ON ojt_trainee_documents;
CREATE POLICY "Allow all on ojt_trainee_documents" ON ojt_trainee_documents FOR ALL USING (true) WITH CHECK (true);

-- 3) OJT Attendance (one row per trainee per day)
CREATE TABLE IF NOT EXISTS ojt_attendance (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    trainee_id UUID NOT NULL REFERENCES ojt_trainees(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    am_in_time TEXT,
    am_in_photo TEXT,
    am_out_time TEXT,
    pm_in_time TEXT,
    pm_in_photo TEXT,
    pm_out_time TEXT,
    ot_in_time TEXT,
    ot_in_photo TEXT,
    ot_out_time TEXT,
    ot_closed BOOLEAN DEFAULT false,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(trainee_id, date)
);
ALTER TABLE ojt_attendance ADD COLUMN IF NOT EXISTS ot_closed BOOLEAN DEFAULT false;

ALTER TABLE ojt_attendance ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_attendance" ON ojt_attendance;
CREATE POLICY "Allow all on ojt_attendance" ON ojt_attendance FOR ALL USING (true) WITH CHECK (true);

-- 3b) Re-open attendance (admin can reopen buttons for a trainee for a given day)
CREATE TABLE IF NOT EXISTS ojt_attendance_reopen (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    trainee_id UUID NOT NULL REFERENCES ojt_trainees(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(trainee_id, date)
);
ALTER TABLE ojt_attendance_reopen ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_attendance_reopen" ON ojt_attendance_reopen;
CREATE POLICY "Allow all on ojt_attendance_reopen" ON ojt_attendance_reopen FOR ALL USING (true) WITH CHECK (true);

-- 4) OJT Admin (optional — default admin for login)
CREATE TABLE IF NOT EXISTS ojt_admin (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE ojt_admin ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow all on ojt_admin" ON ojt_admin;
CREATE POLICY "Allow all on ojt_admin" ON ojt_admin FOR ALL USING (true) WITH CHECK (true);

-- Insert default admin (password admin123 — store as plain for prototype; use hash in production)
INSERT INTO ojt_admin (username, password_hash) VALUES ('admin', 'admin123')
ON CONFLICT (username) DO NOTHING;
