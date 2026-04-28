import { notFound } from "next/navigation";
import { StatusPoller } from "@/components/StatusPoller";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";

type ProcessPageProps = {
  params: Promise<{
    jobId: string;
  }>;
};

export default async function ProcessPage({ params }: ProcessPageProps) {
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
      phase: true,
      progressCurrent: true,
      progressTotal: true,
      errorMessage: true,
    },
  });

  if (!job) {
    notFound();
  }

  return (
    <main className="mx-auto flex min-h-screen max-w-3xl flex-col justify-center gap-6 px-6 py-16">
      <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
        Stage 2
      </p>
      <h1 className="text-4xl font-bold tracking-tight text-slate-950">
        Processing deck
      </h1>
      <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-sm">
        <dl className="space-y-3 text-slate-700">
          <div>
            <dt className="font-semibold text-slate-950">File</dt>
            <dd>{job.uploadedFilename}</dd>
          </div>
          <div>
            <dt className="font-semibold text-slate-950">Status</dt>
            <dd>{job.status}</dd>
          </div>
          <div>
            <dt className="font-semibold text-slate-950">Phase</dt>
            <dd>{job.phase ?? "Queued"}</dd>
          </div>
          <div>
            <dt className="font-semibold text-slate-950">Progress</dt>
            <dd>
              {job.progressCurrent} / {job.progressTotal}
            </dd>
          </div>
          {job.errorMessage ? (
            <div>
              <dt className="font-semibold text-red-700">Error</dt>
              <dd>{job.errorMessage}</dd>
            </div>
          ) : null}
        </dl>
      </div>
      <StatusPoller jobId={jobId} />
    </main>
  );
}
