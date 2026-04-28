import "dotenv/config";
import { defineConfig, env } from "prisma/config";

export default defineConfig({
  schema: "prisma/schema.prisma",
  migrations: {
    path: "prisma/migrations",
  },
  datasource: {
    // Prisma CLI commands use the direct Supabase connection. The app runtime
    // uses DATABASE_URL through the PostgreSQL adapter in src/lib/db.ts.
    url: env("DIRECT_URL"),
  },
});
