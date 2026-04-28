import { NextResponse, type NextRequest } from "next/server";
import { safeInternalRedirectPath } from "@/lib/redirects";
import { createSupabaseMiddlewareClient } from "@/lib/supabase";

export async function GET(request: NextRequest) {
  const requestUrl = new URL(request.url);
  const code = requestUrl.searchParams.get("code");
  const next = safeInternalRedirectPath(requestUrl.searchParams.get("next"));
  let response = NextResponse.redirect(new URL(next, request.url));

  if (code) {
    const supabase = createSupabaseMiddlewareClient(request, response);
    const { error } = await supabase.auth.exchangeCodeForSession(code);

    if (error) {
      response = NextResponse.redirect(
        new URL(`/auth/sign-in?message=${encodeURIComponent(error.message)}`, request.url),
      );
    }
  }

  return response;
}
