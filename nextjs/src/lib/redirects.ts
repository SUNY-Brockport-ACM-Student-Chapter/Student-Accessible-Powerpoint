export function safeInternalRedirectPath(value: string | null | undefined, fallback = "/upload") {
  if (!value || !value.startsWith("/") || value.startsWith("//")) {
    return fallback;
  }

  return value;
}
