import { requireEnv } from "@/lib/env";

type StartJobPayload = {
  storage_object: string;
  presentation_name: string;
};

export async function startProcessingJob(jobId: string, payload: StartJobPayload) {
  const baseUrl = requireEnv("PY_SERVICE_URL").replace(/\/+$/, "");
  const response = await fetch(`${baseUrl}/jobs/${jobId}/start`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-sap-processor-secret": requireEnv("PY_SERVICE_SHARED_SECRET"),
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(`Processor start failed with ${response.status}`);
  }
}
