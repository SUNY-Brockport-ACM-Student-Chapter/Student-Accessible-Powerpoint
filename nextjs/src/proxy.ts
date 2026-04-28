import { NextResponse, type NextRequest } from "next/server";
import { createSupabaseMiddlewareClient } from "@/lib/supabase";

const PUBLIC_PATHS = ["/", "/auth/sign-in", "/auth/callback"];
const CONSENT_PATH = "/consent";
const PROCESSOR_WEBHOOK_PATH = "/api/webhooks/processor";

function isStaticAsset(pathname: string) {
  return (
    pathname.startsWith("/_next/") ||
    pathname === "/favicon.ico" ||
    /\.[a-zA-Z0-9]+$/.test(pathname)
  );
}

function isPublicPath(pathname: string) {
  return (
    PUBLIC_PATHS.includes(pathname) ||
    pathname.startsWith("/auth/") ||
    pathname === PROCESSOR_WEBHOOK_PATH ||
    isStaticAsset(pathname)
  );
}

export async function proxy(request: NextRequest) {
  const pathname = request.nextUrl.pathname;
  let response = NextResponse.next({
    request,
  });

  if (isPublicPath(pathname)) {
    return response;
  }

  const supabase = createSupabaseMiddlewareClient(request, response);
  const {
    data: { user },
  } = await supabase.auth.getUser();

  if (!user) {
    const redirectUrl = request.nextUrl.clone();
    redirectUrl.pathname = "/auth/sign-in";
    redirectUrl.searchParams.set("redirectedFrom", pathname);
    return NextResponse.redirect(redirectUrl);
  }

  if (pathname !== CONSENT_PATH) {
    const { data: profile } = await supabase
      .from("Profile")
      .select("consentAcceptedAt")
      .eq("id", user.id)
      .maybeSingle();

    if (!profile?.consentAcceptedAt) {
      const redirectUrl = request.nextUrl.clone();
      redirectUrl.pathname = CONSENT_PATH;
      redirectUrl.search = "";
      response = NextResponse.redirect(redirectUrl);
    }
  }

  return response;
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
