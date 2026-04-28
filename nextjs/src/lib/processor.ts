import { requireEnv } from "@/lib/env";

type StartJobPayload = {
  storage_object: string;
  presentation_name: string;
};

type CommitJobPayload = {
  storageObject: string;
  presentationName: string;
  descriptions: Array<{
    slideNumber: number;
    orderNumber: number;
    altText: string;
  }>;
};

async function postProcessorJob(jobId: string, action: "start" | "commit", payload: unknown) {
  const baseUrl = requireEnv("PY_SERVICE_URL").replace(/\/+$/, "");
  const response = await fetch(`${baseUrl}/jobs/${jobId}/${action}`, {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-sap-processor-secret": requireEnv("PY_SERVICE_SHARED_SECRET"),
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(`Processor ${action} failed with ${response.status}`);
  }
}

export async function startProcessingJob(jobId: string, payload: StartJobPayload) {
  await postProcessorJob(jobId, "start", payload);
}

export async function commitProcessingJob(jobId: string, payload: CommitJobPayload) {
  await postProcessorJob(jobId, "commit", {
    storage_object: payload.storageObject,
    presentation_name: payload.presentationName,
    descriptions: payload.descriptions.map((description) => ({
      slide_number: description.slideNumber,
      order_number: description.orderNumber,
      alt_text: description.altText,
    })),
  });
}
