import { prisma } from "@/lib/db";
import { createSupabaseServerClient } from "@/lib/supabase";

export async function getCurrentUser() {
  const supabase = await createSupabaseServerClient();
  const {
    data: { user },
    error,
  } = await supabase.auth.getUser();

  if (error || !user) {
    return null;
  }

  return user;
}

export async function requireCurrentProfile() {
  const user = await getCurrentUser();
  if (!user?.email) {
    return null;
  }

  return prisma.profile.upsert({
    where: { id: user.id },
    update: { email: user.email },
    create: {
      id: user.id,
      email: user.email,
    },
  });
}
