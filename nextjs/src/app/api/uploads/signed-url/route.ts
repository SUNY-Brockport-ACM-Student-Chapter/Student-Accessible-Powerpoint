import { fail, ok } from "@/lib/api";
import { requireCurrentProfile } from "@/lib/auth";
import { requireEnv } from "@/lib/env";
import {
  buildUploadObjectPath,
  createPresentationSignedUploadUrl,
  isPptxFilename,
  MAX_UPLOAD_BYTES,
} from "@/lib/storage";

export const runtime = "nodejs";

export async function POST(request: Request) {
  const profile = await requireCurrentProfile();
  if (!profile) {
    return fail(
      {
        code: "UNAUTHORIZED",
        message: "Sign in before uploading a deck.",
        retryable: false,
      },
      401,
    );
  }

  if (!profile.consentAcceptedAt) {
    return fail(
      {
        code: "CONSENT_REQUIRED",
        message: "Accept the consent form before uploading a deck.",
        retryable: false,
      },
      403,
    );
  }

  const body = await request.json().catch(() => null);
  const filename = typeof body?.filename === "string" ? body.filename : "";
  const sizeBytes = typeof body?.sizeBytes === "number" ? body.sizeBytes : 0;

  if (!isPptxFilename(filename)) {
    return fail(
      {
        code: "INVALID_FILE",
        message: "Upload a .pptx PowerPoint file.",
        retryable: false,
      },
      400,
    );
  }

  if (sizeBytes > MAX_UPLOAD_BYTES) {
    return fail(
      {
        code: "UPLOAD_TOO_LARGE",
        message: "PowerPoint uploads are limited to 50 MB.",
        retryable: false,
      },
      413,
    );
  }

  const objectPath = buildUploadObjectPath(profile.id, filename);
  const signedUpload = await createPresentationSignedUploadUrl(objectPath);

  return ok({
    ...signedUpload,
    bucket: requireEnv("SUPABASE_UPLOADS_BUCKET"),
  });
}
