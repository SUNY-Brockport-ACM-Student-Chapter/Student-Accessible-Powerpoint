export type ApiErrorCode =
  | "UPLOAD_TOO_LARGE"
  | "INVALID_FILE"
  | "CONSENT_REQUIRED"
  | "PROCESSOR_UNAVAILABLE"
  | "JOB_NOT_FOUND"
  | "UNAUTHORIZED"
  | "UNKNOWN";

export type ApiError = {
  code: ApiErrorCode;
  message: string;
  retryable: boolean;
};

export function ok<T>(data: T) {
  return Response.json({ ok: true, data });
}

export function fail(error: ApiError, status: number) {
  return Response.json({ ok: false, error }, { status });
}
