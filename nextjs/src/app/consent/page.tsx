import { notFound, redirect } from "next/navigation";
import { acceptConsent, CONSENT_VERSION, getConsentMarkdown } from "@/lib/consent";
import { requireCurrentProfile } from "@/lib/auth";

function markdownToParagraphs(markdown: string) {
  return markdown
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean);
}

export default async function ConsentPage() {
  const profile = await requireCurrentProfile();
  if (!profile) {
    notFound();
  }

  if (profile.consentAcceptedAt && profile.consentVersion === CONSENT_VERSION) {
    redirect("/upload");
  }

  const consentMarkdown = await getConsentMarkdown();
  const blocks = markdownToParagraphs(consentMarkdown);

  return (
    <main className="mx-auto flex min-h-screen max-w-3xl flex-col gap-8 px-6 py-16">
      <div className="space-y-3">
        <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
          Consent
        </p>
        <h1 className="text-4xl font-bold tracking-tight text-slate-950">
          Review consent before uploading
        </h1>
        <p className="text-lg leading-8 text-slate-700">
          Consent version: `{CONSENT_VERSION}`
        </p>
      </div>

      <article className="space-y-4 rounded-2xl border border-slate-200 bg-white p-6 text-slate-700 shadow-sm">
        {blocks.map((block) => (
          <p className="leading-7" key={block}>
            {block.replace(/^#+\s*/, "").replace(/^>\s*/, "")}
          </p>
        ))}
      </article>

      <form action={acceptConsent}>
        <button className="rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white">
          I consent and want to continue
        </button>
      </form>
    </main>
  );
}
