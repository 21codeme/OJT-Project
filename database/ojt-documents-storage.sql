-- OJT trainee documents: Storage bucket policies (prototype — public read/write)
-- 1) In Supabase Dashboard → Storage → New bucket:
--    Name: ojt-trainee-documents
--    Public bucket: ON (for direct <img> / <iframe> preview without signed URLs)
-- 2) Run this entire script in the SQL Editor.

DROP POLICY IF EXISTS "Allow public uploads to ojt-trainee-documents" ON storage.objects;
DROP POLICY IF EXISTS "Allow public read from ojt-trainee-documents" ON storage.objects;
DROP POLICY IF EXISTS "Allow public update in ojt-trainee-documents" ON storage.objects;
DROP POLICY IF EXISTS "Allow public delete from ojt-trainee-documents" ON storage.objects;

CREATE POLICY "Allow public uploads to ojt-trainee-documents"
ON storage.objects FOR INSERT
TO public
WITH CHECK (bucket_id = 'ojt-trainee-documents');

CREATE POLICY "Allow public read from ojt-trainee-documents"
ON storage.objects FOR SELECT
TO public
USING (bucket_id = 'ojt-trainee-documents');

CREATE POLICY "Allow public update in ojt-trainee-documents"
ON storage.objects FOR UPDATE
TO public
USING (bucket_id = 'ojt-trainee-documents')
WITH CHECK (bucket_id = 'ojt-trainee-documents');

CREATE POLICY "Allow public delete from ojt-trainee-documents"
ON storage.objects FOR DELETE
TO public
USING (bucket_id = 'ojt-trainee-documents');
