"use client";

import { useRouter } from "next/navigation";
import { useEffect, useState } from "react";

type JobStatusPayload = {
  status: string;
  phase: string | null;
  progressCurrent: number;
  progressTotal: number;
  errorMessage: string | null;
};

export function StatusPoller({ jobId }: { jobId: string }) {
  const router = useRouter();
  const [job, setJob] = useState<JobStatusPayload | null>(null);

  useEffect(() => {
    let cancelled = false;
    let timeoutId: number | undefined;
    const startedAt = Date.now();

    async function poll() {
      const response = await fetch(`/api/jobs/${jobId}`, { cache: "no-store" });
      const payload = await response.json();
      if (cancelled || !payload.ok) {
        return;
      }

      const nextJob = payload.data.job as JobStatusPayload;
      setJob(nextJob);

      if (nextJob.status === "awaiting_review") {
        router.push(`/review/${jobId}`);
      } else if (nextJob.status === "ready") {
        router.push(`/download/${jobId}`);
      } else if (nextJob.status !== "error") {
        const elapsedMs = Date.now() - startedAt;
        timeoutId = window.setTimeout(() => void poll(), elapsedMs > 120000 ? 10000 : 5000);
      }
    }

    void poll();

    return () => {
      cancelled = true;
      if (timeoutId !== undefined) {
        window.clearTimeout(timeoutId);
      }
    };
  }, [jobId, router]);

  if (!job) {
    return <p className="text-slate-600">Checking current job status...</p>;
  }

  if (job.status === "error") {
    return (
      <p className="text-red-700">
        Processing failed: {job.errorMessage ?? "Unknown error"}
      </p>
    );
  }

  return (
    <p className="text-slate-600">
      {job.phase ?? job.status} ({job.progressCurrent} / {job.progressTotal})
    </p>
  );
}
