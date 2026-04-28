import { JobStatus } from "@/generated/prisma/enums";
import { fail, ok } from "@/lib/api";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { startProcessingJob } from "@/lib/processor";
import {
  buildUploadObjectPath,
  isPptxFile,
  MAX_UPLOAD_BYTES,
  uploadPresentationToStorage,
} from "@/lib/storage";

export const runtime = "nodejs";

export async function POST(request: Request) {
  const contentLength = Number(request.headers.get("content-length") ?? 0);
  if (contentLength > MAX_UPLOAD_BYTES) {
    return fail(
      {
        code: "UPLOAD_TOO_LARGE",
        message: "PowerPoint uploads are limited to 50 MB.",
        retryable: false,
      },
      413,
    );
  }

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

  const formData = await request.formData();
  const file = formData.get("file");

  if (!(file instanceof File) || !isPptxFile(file)) {
    return fail(
      {
        code: "INVALID_FILE",
        message: "Upload a .pptx PowerPoint file.",
        retryable: false,
      },
      400,
    );
  }

  if (file.size > MAX_UPLOAD_BYTES) {
    return fail(
      {
        code: "UPLOAD_TOO_LARGE",
        message: "PowerPoint uploads are limited to 50 MB.",
        retryable: false,
      },
      413,
    );
  }

  const objectPath = buildUploadObjectPath(profile.id, file.name);
  const storagePath = await uploadPresentationToStorage({ file, objectPath });

  const job = await prisma.job.create({
    data: {
      profileId: profile.id,
      uploadedFilename: file.name,
      uploadObjectPath: storagePath,
      status: JobStatus.queued,
    },
    select: {
      id: true,
    },
  });

  try {
    await startProcessingJob(job.id, {
      storage_object: storagePath,
      presentation_name: file.name,
    });
  } catch (error) {
    await prisma.job.update({
      where: { id: job.id },
      data: {
        status: JobStatus.error,
        errorCode: "PROCESSOR_UNAVAILABLE",
        errorMessage: error instanceof Error ? error.message : "Processor unavailable",
      },
    });

    return fail(
      {
        code: "PROCESSOR_UNAVAILABLE",
        message: "The processing service did not accept the upload.",
        retryable: true,
      },
      502,
    );
  }

  return ok({ jobId: job.id });
}
