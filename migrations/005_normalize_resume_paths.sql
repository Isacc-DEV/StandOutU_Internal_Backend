-- Normalize legacy resume paths to the new /resumes/<file> convention.
UPDATE resumes
SET file_path = REPLACE(file_path, '/data/resumes/', '/resumes/')
WHERE file_path LIKE '/data/resumes/%';

-- Windows-style legacy paths (optional). This will strip drive letter; adjust if needed.
UPDATE resumes
SET file_path = '/resumes/' || regexp_replace(file_path, '^.*data[\\\\/]resumes[\\\\/]', '')
WHERE file_path ~* 'data[\\\\/]resumes[\\\\/]';
