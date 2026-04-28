import { notFound, redirect } from "next/navigation";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";

type DownloadPageProps = {
  params: Promise<{
    jobId: string;
  }>;
};

export default async function DownloadPage({ params }: DownloadPageProps) {
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
      outputObjectPath: true,
    },
  });

  if (!job) {
    notFound();
  }

  if (job.status !== "ready") {
    redirect(`/process/${jobId}`);
  }

  return (
    <main className="mx-auto flex min-h-screen max-w-3xl flex-col justify-center gap-6 px-6 py-16">
      <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
        Stage 4
      </p>
      <h1 className="text-4xl font-bold tracking-tight text-slate-950">
        Accessible deck ready
      </h1>
      <p className="text-lg leading-8 text-slate-700">
        `{job.uploadedFilename}` has been rebuilt. The signed download button
        will be added in the download-flow slice.
      </p>
      <p className="rounded-lg bg-slate-100 p-4 font-mono text-sm text-slate-700">
        {job.outputObjectPath}
      </p>
    </main>
  );
}
