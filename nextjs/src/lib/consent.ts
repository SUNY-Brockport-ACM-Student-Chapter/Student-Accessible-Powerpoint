import { createHash } from "crypto";
import { readFile } from "fs/promises";
import path from "path";
import { headers } from "next/headers";
import { redirect } from "next/navigation";
import { requireCurrentProfile } from "@/lib/auth";
import { prisma } from "@/lib/db";

export const CONSENT_VERSION = "consent-v1";

export async function getConsentMarkdown() {
  return readFile(
    path.join(process.cwd(), "src", "content", `${CONSENT_VERSION}.md`),
    "utf8",
  );
}

function hashIp(ip: string) {
  return createHash("sha256").update(ip).digest("hex");
}

export async function acceptConsent() {
  "use server";

  const profile = await requireCurrentProfile();
  if (!profile) {
    redirect("/auth/sign-in");
  }

  const headerStore = await headers();
  const forwardedFor = headerStore.get("x-forwarded-for") ?? "";
  const ip = forwardedFor.split(",")[0]?.trim() || "unknown";
  const userAgent = headerStore.get("user-agent");
  const acceptedAt = new Date();

  await prisma.$transaction([
    prisma.consentEvent.create({
      data: {
        profileId: profile.id,
        acceptedAt,
        consentVersion: CONSENT_VERSION,
        ipHash: hashIp(ip),
        userAgent,
      },
    }),
    prisma.profile.update({
      where: { id: profile.id },
      data: {
        consentAcceptedAt: acceptedAt,
        consentVersion: CONSENT_VERSION,
      },
    }),
  ]);

  redirect("/upload");
}
