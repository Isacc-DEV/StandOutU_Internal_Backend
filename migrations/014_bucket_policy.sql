CREATE POLICY "Allow public users to upload community-files"
ON storage.objects FOR INSERT
TO public
WITH CHECK (bucket_id = 'community-files');

CREATE POLICY "Allow public to view community-files"
ON storage.objects FOR SELECT
TO public
USING (bucket_id = 'community-files');
