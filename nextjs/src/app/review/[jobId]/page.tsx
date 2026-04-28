import { notFound, redirect } from "next/navigation";
import { ReviewGrid } from "@/components/ReviewGrid";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";

type ReviewPageProps = {
  params: Promise<{
    jobId: string;
  }>;
};

export default async function ReviewPage({ params }: ReviewPageProps) {
  const profile = await requireCurrentProfile();
  if (!profile) {
    notFound();
  }

  const { jobId } = await params;
  const job = await prisma.job.findFirst({
    where: {
      id: jobId,
      profileId: profile.id,
    },
    select: {
      uploadedFilename: true,
      status: true,
      descriptions: {
        orderBy: [{ slideNumber: "asc" }, { orderNumber: "asc" }],
        select: {
          id: true,
          slideNumber: true,
          orderNumber: true,
          aiDescription: true,
          finalAltText: true,
        },
      },
    },
  });

  if (!job) {
    notFound();
  }

  if (job.status === "ready") {
    redirect(`/download/${jobId}`);
  }

  if (job.status !== "awaiting_review") {
    redirect(`/process/${jobId}`);
  }

  return (
    <main className="mx-auto flex min-h-screen max-w-4xl flex-col gap-8 px-6 py-16">
      <div className="space-y-3">
        <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
          Stage 3
        </p>
        <h1 className="text-4xl font-bold tracking-tight text-slate-950">
          Review generated alt text
        </h1>
        <p className="text-lg leading-8 text-slate-700">
          Review each image description for `{job.uploadedFilename}` before
          exporting the accessible deck.
        </p>
      </div>

      <ReviewGrid jobId={jobId} descriptions={job.descriptions} />
    </main>
  );
}
