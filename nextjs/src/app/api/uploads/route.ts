import { JobStatus } from "@/generated/prisma/enums";
import { fail, ok } from "@/lib/api";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { startProcessingJob } from "@/lib/processor";
import { isPptxFilename, verifyPresentationUploadExists } from "@/lib/storage";

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
  const storagePath = typeof body?.storageObject === "string" ? body.storageObject : "";
  const presentationName =
    typeof body?.presentationName === "string" ? body.presentationName : "";

  if (!storagePath.startsWith(`${profile.id}/`) || !isPptxFilename(presentationName)) {
    return fail(
      {
        code: "INVALID_FILE",
        message: "Complete a signed .pptx upload before creating a job.",
        retryable: false,
      },
      400,
    );
  }

  const uploadExists = await verifyPresentationUploadExists(storagePath);
  if (!uploadExists) {
    return fail(
      {
        code: "INVALID_FILE",
        message: "Uploaded deck was not found in storage.",
        retryable: true,
      },
      400,
    );
  }

  const job = await prisma.job.create({
    data: {
      profileId: profile.id,
      uploadedFilename: presentationName,
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
      presentation_name: presentationName,
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
