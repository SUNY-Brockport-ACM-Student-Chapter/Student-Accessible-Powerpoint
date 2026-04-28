const SERVER_ONLY_ENV = [
  "DATABASE_URL",
  "DIRECT_URL",
  "SUPABASE_SERVICE_ROLE_KEY",
  "PY_SERVICE_SHARED_SECRET",
] as const;

export function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export function assertServerOnlyEnvIsNotPublic() {
  for (const name of SERVER_ONLY_ENV) {
    if (name.startsWith("NEXT_PUBLIC_")) {
      throw new Error(`Server-only environment variable is public: ${name}`);
    }
  }
}
