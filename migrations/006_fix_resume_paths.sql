-- Normalize resume file paths to /data/resumes/<file>
UPDATE resumes
SET file_path = REPLACE(file_path, '/resumes/', '/data/resumes/')
WHERE file_path LIKE '/resumes/%';
