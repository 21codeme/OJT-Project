-- Supabase Storage policies for bucket: ojt-daily-logs
-- Create the bucket first in Storage UI (Public: ON), then run this SQL.

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
