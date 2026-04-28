import { notFound, redirect } from "next/navigation";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";
import { createOutputSignedDownloadUrl } from "@/lib/storage";

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

  if (!job.outputObjectPath) {
    notFound();
  }

  const signedUrl = await createOutputSignedDownloadUrl(job.outputObjectPath);

  return (
    <main className="mx-auto flex min-h-screen max-w-3xl flex-col justify-center gap-6 px-6 py-16">
      <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
        Stage 4
      </p>
      <h1 className="text-4xl font-bold tracking-tight text-slate-950">
        Accessible deck ready
      </h1>
      <p className="text-lg leading-8 text-slate-700">
        `{job.uploadedFilename}` has been rebuilt. This link expires in 10
        minutes.
      </p>
      <a
        className="w-fit rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white"
        href={signedUrl}
      >
        Download Accessible PowerPoint
      </a>
    </main>
  );
}
