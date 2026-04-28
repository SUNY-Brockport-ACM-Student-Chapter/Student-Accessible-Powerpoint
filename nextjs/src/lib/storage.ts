import { createClient } from "@supabase/supabase-js";
import { MAX_UPLOAD_BYTES, PPTX_MIME_TYPE } from "@/lib/constants";
import { requireEnv } from "@/lib/env";

export { MAX_UPLOAD_BYTES, PPTX_MIME_TYPE };

function createSupabaseAdminClient() {
  return createClient(
    requireEnv("NEXT_PUBLIC_SUPABASE_URL"),
    requireEnv("SUPABASE_SERVICE_ROLE_KEY"),
    {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    },
  );
}

function sanitizeFilename(filename: string) {
  return filename.replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 160);
}

export function isPptxFilename(filename: string) {
  return filename.toLowerCase().endsWith(".pptx");
}

export function isPptxFile(file: File) {
  return isPptxFilename(file.name);
}

export function buildUploadObjectPath(profileId: string, filename: string) {
  return `${profileId}/${crypto.randomUUID()}-${sanitizeFilename(filename)}`;
}

export async function createPresentationSignedUploadUrl(objectPath: string) {
  const supabase = createSupabaseAdminClient();
  const bucket = requireEnv("SUPABASE_UPLOADS_BUCKET");

  const { data, error } = await supabase.storage
    .from(bucket)
    .createSignedUploadUrl(objectPath);

  if (error) {
    throw error;
  }

  return {
    path: data.path,
    token: data.token,
  };
}

export async function verifyPresentationUploadExists(objectPath: string) {
  const supabase = createSupabaseAdminClient();
  const bucket = requireEnv("SUPABASE_UPLOADS_BUCKET");
  const lastSlash = objectPath.lastIndexOf("/");
  const folder = lastSlash === -1 ? "" : objectPath.slice(0, lastSlash);
  const filename = lastSlash === -1 ? objectPath : objectPath.slice(lastSlash + 1);

  const { data, error } = await supabase.storage.from(bucket).list(folder, {
    limit: 100,
    search: filename,
  });

  if (error) {
    throw error;
  }

  return data.some((object) => object.name === filename);
}
