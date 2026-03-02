-- Backup tables sa Supabase — hiwalay na table para sa Inventory backup at Class Schedule backup
-- I-run sa Supabase SQL Editor (New Query > Paste > Run)

-- 1) Inventory Backup Snapshot — isang row (id=1), nakastore ang buong backup JSON
CREATE TABLE IF NOT EXISTS inventory_backup_snapshot (
    id INTEGER PRIMARY KEY DEFAULT 1,
    data JSONB NOT NULL DEFAULT '{}',
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    CONSTRAINT single_row CHECK (id = 1)
);

-- I-insert ang default row kung wala pa
INSERT INTO inventory_backup_snapshot (id, data, updated_at)
VALUES (1, '{}', NOW())
ON CONFLICT (id) DO NOTHING;

ALTER TABLE inventory_backup_snapshot ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Allow all on inventory_backup_snapshot" ON inventory_backup_snapshot
    FOR ALL USING (true) WITH CHECK (true);

CREATE OR REPLACE FUNCTION update_inventory_backup_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;
CREATE TRIGGER inventory_backup_snapshot_updated_at
    BEFORE UPDATE ON inventory_backup_snapshot
    FOR EACH ROW EXECUTE FUNCTION update_inventory_backup_updated_at();


-- 2) Class Schedule Backup Snapshot — isang row (id=1), nakastore ang buong backup JSON
CREATE TABLE IF NOT EXISTS class_schedule_backup_snapshot (
    id INTEGER PRIMARY KEY DEFAULT 1,
    data JSONB NOT NULL DEFAULT '{}',
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    CONSTRAINT single_row_schedule CHECK (id = 1)
);

INSERT INTO class_schedule_backup_snapshot (id, data, updated_at)
VALUES (1, '{}', NOW())
ON CONFLICT (id) DO NOTHING;

ALTER TABLE class_schedule_backup_snapshot ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Allow all on class_schedule_backup_snapshot" ON class_schedule_backup_snapshot
    FOR ALL USING (true) WITH CHECK (true);

CREATE OR REPLACE FUNCTION update_schedule_backup_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;
CREATE TRIGGER class_schedule_backup_snapshot_updated_at
    BEFORE UPDATE ON class_schedule_backup_snapshot
    FOR EACH ROW EXECUTE FUNCTION update_schedule_backup_updated_at();
