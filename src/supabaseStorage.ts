import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_PUBLISHABLE_KEY;
const BUCKET_NAME = process.env.COMMUNITY_FILES_BUCKET_STORAGE || 'community-files';

const supabase = SUPABASE_URL && SUPABASE_KEY ? createClient(SUPABASE_URL, SUPABASE_KEY) : null;

export async function uploadFile(
  file: Buffer,
  fileName: string,
  mimeType: string,
): Promise<{ url: string; path: string }> {
  if (!supabase) {
    throw new Error('Supabase is not configured');
  }
  const timestamp = Date.now();
  const safeName = fileName.replace(/[^a-zA-Z0-9.-]/g, '_');
  const filePath = `${timestamp}_${safeName}`;

  const { data, error } = await supabase.storage
    .from(BUCKET_NAME)
    .upload(filePath, file, {
      contentType: mimeType,
      upsert: false,
    });

  if (error) {
    throw new Error(`Supabase upload failed: ${error.message}`);
  }

  const { data: urlData } = supabase.storage.from(BUCKET_NAME).getPublicUrl(filePath);

  return {
    url: urlData.publicUrl,
    path: filePath,
  };
}

export async function deleteFile(filePath: string): Promise<void> {
  if (!supabase) {
    throw new Error('Supabase is not configured');
  }
  const { error } = await supabase.storage.from(BUCKET_NAME).remove([filePath]);
  if (error) {
    console.error('Supabase delete failed:', error);
  }
}
