-- Supabase Storage: Allow upload and read for bucket "inventory-pictures"
-- Run this in Supabase SQL Editor after creating the bucket (Dashboard → Storage → New bucket: inventory-pictures, Public: ON)
-- Fixes: "new row violates row-level security policy" when uploading pictures

-- Allow anyone (including anon) to INSERT (upload) into inventory-pictures
CREATE POLICY "Allow public uploads to inventory-pictures"
ON storage.objects FOR INSERT
TO public
WITH CHECK (bucket_id = 'inventory-pictures');

-- Allow anyone to SELECT (read) from inventory-pictures (for public image URLs)
CREATE POLICY "Allow public read from inventory-pictures"
ON storage.objects FOR SELECT
TO public
USING (bucket_id = 'inventory-pictures');

-- Optional: allow update/delete so app can replace or remove images later
CREATE POLICY "Allow public update in inventory-pictures"
ON storage.objects FOR UPDATE
TO public
USING (bucket_id = 'inventory-pictures')
WITH CHECK (bucket_id = 'inventory-pictures');

CREATE POLICY "Allow public delete from inventory-pictures"
ON storage.objects FOR DELETE
TO public
USING (bucket_id = 'inventory-pictures');
