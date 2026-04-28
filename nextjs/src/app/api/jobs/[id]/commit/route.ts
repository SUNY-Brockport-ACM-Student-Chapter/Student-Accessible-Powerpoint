import { JobStatus } from "@/generated/prisma/enums";
import { fail, ok } from "@/lib/api";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { commitProcessingJob } from "@/lib/processor";

export const runtime = "nodejs";

type CommitDescriptionInput = {
  id?: string;
  finalAltText?: string;
};

export async function POST(
  request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const profile = await requireCurrentProfile();
  if (!profile) {
    return fail(
      {
        code: "UNAUTHORIZED",
        message: "Sign in before committing descriptions.",
        retryable: false,
      },
      401,
    );
  }

  const { id } = await context.params;
  const body = await request.json().catch(() => null);
  const inputDescriptions = Array.isArray(body?.descriptions)
    ? (body.descriptions as CommitDescriptionInput[])
    : [];

  const job = await prisma.job.findFirst({
    where: {
      id,
      profileId: profile.id,
    },
    select: {
      id: true,
      uploadedFilename: true,
      uploadObjectPath: true,
      descriptions: true,
    },
  });

  if (!job) {
    return fail(
      {
        code: "JOB_NOT_FOUND",
        message: "Job not found.",
        retryable: false,
      },
      404,
    );
  }

  const submittedTextById = new Map(
    inputDescriptions
      .filter((description) => description.id && description.finalAltText)
      .map((description) => [description.id as string, description.finalAltText as string]),
  );

  const commitDescriptions = job.descriptions.map((description) => ({
    id: description.id,
    slideNumber: description.slideNumber,
    orderNumber: description.orderNumber,
    altText:
      submittedTextById.get(description.id)?.trim() ??
      description.finalAltText ??
      description.aiDescription ??
      "",
  }));

  await prisma.$transaction([
    ...commitDescriptions.map((description) =>
      prisma.slideDescription.update({
        where: { id: description.id },
        data: { finalAltText: description.altText },
      }),
    ),
    prisma.job.update({
      where: { id: job.id },
      data: {
        status: JobStatus.rebuilding,
        phase: "Rebuilding deck",
        committedAt: new Date(),
      },
    }),
  ]);

  try {
    await commitProcessingJob(job.id, {
      storageObject: job.uploadObjectPath,
      presentationName: job.uploadedFilename,
      descriptions: commitDescriptions,
    });
  } catch (error) {
    await prisma.job.update({
      where: { id: job.id },
      data: {
        status: JobStatus.error,
        errorCode: "REBUILD_FAILED",
        errorMessage: error instanceof Error ? error.message : "Processor commit failed",
      },
    });

    return fail(
      {
        code: "PROCESSOR_UNAVAILABLE",
        message: "The processing service did not accept the rebuild request.",
        retryable: true,
      },
      502,
    );
  }

  return ok({ jobId: job.id });
}
