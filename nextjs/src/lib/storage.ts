import { createClient } from "@supabase/supabase-js";
import { requireEnv } from "@/lib/env";

export const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;
export const PPTX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function createSupabaseAdminClient() {
  return createClient(requireEnv("NEXT_PUBLIC_SUPABASE_URL"), requireEnv("SUPABASE_SERVICE_ROLE_KEY"), {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
    },
  });
}

function sanitizeFilename(filename: string) {
  return filename.replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 160);
}

export function isPptxFile(file: File) {
  return file.name.toLowerCase().endsWith(".pptx");
}

export function buildUploadObjectPath(profileId: string, filename: string) {
  return `${profileId}/${crypto.randomUUID()}-${sanitizeFilename(filename)}`;
}

export async function uploadPresentationToStorage(params: {
  file: File;
  objectPath: string;
}) {
  const supabase = createSupabaseAdminClient();
  const bucket = requireEnv("SUPABASE_UPLOADS_BUCKET");

  const { data, error } = await supabase.storage
    .from(bucket)
    .upload(params.objectPath, params.file, {
      contentType: params.file.type || PPTX_MIME_TYPE,
      upsert: false,
    });

  if (error) {
    throw error;
  }

  return data.path;
}
