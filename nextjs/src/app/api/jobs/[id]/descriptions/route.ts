import { fail, ok } from "@/lib/api";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";

export const runtime = "nodejs";

export async function GET(
  _request: Request,
  context: { params: Promise<{ id: string }> },
) {
  const profile = await requireCurrentProfile();
  if (!profile) {
    return fail(
      {
        code: "UNAUTHORIZED",
        message: "Sign in before viewing descriptions.",
        retryable: false,
      },
      401,
    );
  }

  const { id } = await context.params;
  const job = await prisma.job.findFirst({
    where: {
      id,
      profileId: profile.id,
    },
    select: {
      descriptions: {
        orderBy: [{ slideNumber: "asc" }, { orderNumber: "asc" }],
      },
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

  return ok({ descriptions: job.descriptions });
}
