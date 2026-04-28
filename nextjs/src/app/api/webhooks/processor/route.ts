import { JobStatus } from "@/generated/prisma/enums";
import { fail, ok } from "@/lib/api";
import { requireEnv } from "@/lib/env";
import { prisma } from "@/lib/db";

export const runtime = "nodejs";

type ProcessorDescription = {
  slide_number?: number;
  slideNumber?: number;
  order_number?: number;
  orderNumber?: number;
  alt_text?: string;
  altText?: string;
  item_type?: string;
  itemType?: string;
};

function normalizeStatus(status: unknown) {
  if (typeof status !== "string") {
    return null;
  }

  return Object.values(JobStatus).includes(status as JobStatus)
    ? (status as JobStatus)
    : null;
}

function optionalNumber(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function optionalString(value: unknown) {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

export async function POST(request: Request) {
  const secret = request.headers.get("x-sap-processor-secret");
  if (secret !== requireEnv("PY_SERVICE_SHARED_SECRET")) {
    return fail(
      {
        code: "UNAUTHORIZED",
        message: "Processor webhook secret is invalid.",
        retryable: false,
      },
      401,
    );
  }

  const body = await request.json().catch(() => null);
  const jobId = optionalString(body?.jobId ?? body?.job_id);
  const status = normalizeStatus(body?.status);

  if (!jobId || !status) {
    return fail(
      {
        code: "UNKNOWN",
        message: "Processor webhook payload is invalid.",
        retryable: false,
      },
      400,
    );
  }

  const descriptions = Array.isArray(body?.descriptions)
    ? (body.descriptions as ProcessorDescription[])
    : [];
  const now = new Date();

  await prisma.$transaction(async (tx) => {
    await tx.job.update({
      where: { id: jobId },
      data: {
        status,
        phase: optionalString(body?.phase) ?? null,
        progressCurrent: optionalNumber(body?.progressCurrent ?? body?.progress_current) ?? 0,
        progressTotal: optionalNumber(body?.progressTotal ?? body?.progress_total) ?? 0,
        collectionId: optionalString(body?.collectionId ?? body?.collection_id),
        outputObjectPath: optionalString(body?.outputObjectPath ?? body?.output_object_path),
        errorCode: optionalString(body?.errorCode ?? body?.error_code),
        errorMessage: optionalString(body?.errorMessage ?? body?.error_message),
        startedAt: status === JobStatus.parsing ? now : undefined,
        awaitingReviewAt: status === JobStatus.awaiting_review ? now : undefined,
        readyAt: status === JobStatus.ready ? now : undefined,
      },
    });

    for (const description of descriptions) {
      const slideNumber = optionalNumber(description.slideNumber ?? description.slide_number);
      const orderNumber = optionalNumber(description.orderNumber ?? description.order_number);
      const aiDescription = optionalString(description.altText ?? description.alt_text);

      if (slideNumber === undefined || orderNumber === undefined || !aiDescription) {
        continue;
      }

      await tx.slideDescription.upsert({
        where: {
          jobId_slideNumber_orderNumber: {
            jobId,
            slideNumber,
            orderNumber,
          },
        },
        update: {
          aiDescription,
        },
        create: {
          jobId,
          slideNumber,
          orderNumber,
          itemType: optionalString(description.itemType ?? description.item_type) ?? "image",
          aiDescription,
        },
      });
    }
  });

  return ok({ jobId });
}
