import { redirect } from "next/navigation";
import { safeInternalRedirectPath } from "@/lib/redirects";
import { createSupabaseServerClient } from "@/lib/supabase";

type SignInPageProps = {
  searchParams: Promise<{
    redirectedFrom?: string;
    message?: string;
  }>;
};

async function signIn(formData: FormData) {
  "use server";

  const email = String(formData.get("email") ?? "");
  const redirectedFrom = safeInternalRedirectPath(
    String(formData.get("redirectedFrom") ?? "/upload"),
  );
  const supabase = await createSupabaseServerClient();
  const appUrl = process.env.NEXT_PUBLIC_APP_URL ?? "http://localhost:3000";
  const callbackUrl = new URL("/auth/callback", appUrl);
  callbackUrl.searchParams.set("next", redirectedFrom);

  const { error } = await supabase.auth.signInWithOtp({
    email,
    options: {
      emailRedirectTo: callbackUrl.toString(),
    },
  });

  if (error) {
    redirect(`/auth/sign-in?message=${encodeURIComponent(error.message)}`);
  }

  redirect("/auth/sign-in?message=Check your email for the sign-in link.");
}

export default async function SignInPage({ searchParams }: SignInPageProps) {
  const params = await searchParams;

  return (
    <main className="mx-auto flex min-h-screen max-w-xl flex-col justify-center gap-8 px-6 py-16">
      <div className="space-y-3">
        <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
          Sign in
        </p>
        <h1 className="text-4xl font-bold tracking-tight text-slate-950">
          Continue with email
        </h1>
        <p className="text-lg leading-8 text-slate-700">
          Enter your email address and Supabase will send a magic link.
        </p>
      </div>

      <form action={signIn} className="space-y-4 rounded-2xl border border-slate-200 p-6">
        <input
          name="redirectedFrom"
          type="hidden"
          value={safeInternalRedirectPath(params.redirectedFrom)}
        />
        <label className="block space-y-2">
          <span className="font-semibold text-slate-950">Email</span>
          <input
            className="w-full rounded-lg border border-slate-300 p-3"
            name="email"
            required
            type="email"
          />
        </label>
        <button className="rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white">
          Send magic link
        </button>
        {params.message ? <p className="text-sm text-slate-700">{params.message}</p> : null}
      </form>
    </main>
  );
}
