"use client";

import { useRouter } from "next/navigation";
import { useState } from "react";

type ReviewDescription = {
  id: string;
  slideNumber: number;
  orderNumber: number;
  aiDescription: string | null;
  finalAltText: string | null;
};

export function ReviewGrid({
  jobId,
  descriptions,
}: {
  jobId: string;
  descriptions: ReviewDescription[];
}) {
  const router = useRouter();
  const [drafts, setDrafts] = useState(
    () =>
      new Map(
        descriptions.map((description) => [
          description.id,
          description.finalAltText ?? description.aiDescription ?? "",
        ]),
      ),
  );
  const [error, setError] = useState<string | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  async function commitReview() {
    setIsSubmitting(true);
    setError(null);

    const response = await fetch(`/api/jobs/${jobId}/commit`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        descriptions: Array.from(drafts.entries()).map(([id, finalAltText]) => ({
          id,
          finalAltText,
        })),
      }),
    });
    const payload = await response.json();

    if (!response.ok || !payload.ok) {
      setError(payload.error?.message ?? "Review could not be committed.");
      setIsSubmitting(false);
      return;
    }

    router.push(`/process/${jobId}`);
  }

  return (
    <section className="space-y-6">
      {descriptions.map((description) => (
        <article
          className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"
          key={description.id}
        >
          <div className="mb-3 flex flex-wrap gap-3 text-sm font-semibold text-slate-600">
            <span>Slide {description.slideNumber}</span>
            <span>Order {description.orderNumber}</span>
          </div>
          <label className="block space-y-2">
            <span className="font-semibold text-slate-950">Alt text</span>
            <textarea
              className="min-h-32 w-full rounded-lg border border-slate-300 p-3 text-slate-900"
              value={drafts.get(description.id) ?? ""}
              onChange={(event) => {
                const nextDrafts = new Map(drafts);
                nextDrafts.set(description.id, event.target.value);
                setDrafts(nextDrafts);
              }}
            />
          </label>
        </article>
      ))}

      {error ? <p className="text-red-700">{error}</p> : null}

      <button
        className="rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white disabled:opacity-60"
        disabled={isSubmitting || descriptions.length === 0}
        type="button"
        onClick={() => void commitReview()}
      >
        Confirm & Export
      </button>
    </section>
  );
}
