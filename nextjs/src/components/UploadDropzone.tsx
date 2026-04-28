"use client";

import { useRouter } from "next/navigation";
import { useRef, useState } from "react";
import { MAX_UPLOAD_BYTES } from "@/lib/constants";
import { createSupabaseBrowserClient } from "@/lib/supabase-browser";

type UploadState =
  | { status: "idle"; message: string }
  | { status: "uploading"; message: string }
  | { status: "error"; message: string };

export function UploadDropzone() {
  const router = useRouter();
  const inputRef = useRef<HTMLInputElement>(null);
  const [state, setState] = useState<UploadState>({
    status: "idle",
    message: "Drag a .pptx file here, or choose a file.",
  });

  async function submitFile(file: File | undefined) {
    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith(".pptx")) {
      setState({ status: "error", message: "Choose a .pptx PowerPoint file." });
      return;
    }

    if (file.size > MAX_UPLOAD_BYTES) {
      setState({ status: "error", message: "PowerPoint uploads are limited to 50 MB." });
      return;
    }

    setState({ status: "uploading", message: "Preparing private upload..." });

    const signedResponse = await fetch("/api/uploads/signed-url", {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        filename: file.name,
        sizeBytes: file.size,
        contentType: file.type,
      }),
    });
    const signedPayload = await signedResponse.json();

    if (!signedResponse.ok || !signedPayload.ok) {
      setState({
        status: "error",
        message: signedPayload.error?.message ?? "Upload could not be prepared.",
      });
      return;
    }

    setState({ status: "uploading", message: "Uploading deck..." });

    const supabase = createSupabaseBrowserClient();
    const { error: uploadError } = await supabase.storage
      .from(signedPayload.data.bucket)
      .uploadToSignedUrl(signedPayload.data.path, signedPayload.data.token, file);

    if (uploadError) {
      setState({ status: "error", message: uploadError.message });
      return;
    }

    setState({ status: "uploading", message: "Starting processing job..." });

    const response = await fetch("/api/uploads", {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        storageObject: signedPayload.data.path,
        presentationName: file.name,
      }),
    });
    const payload = await response.json();

    if (!response.ok || !payload.ok) {
      setState({
        status: "error",
        message: payload.error?.message ?? "Upload failed.",
      });
      return;
    }

    router.push(`/process/${payload.data.jobId}`);
  }

  return (
    <section
      className="rounded-2xl border border-dashed border-slate-300 bg-slate-50 p-8 text-center"
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => {
        event.preventDefault();
        void submitFile(event.dataTransfer.files[0]);
      }}
    >
      <input
        ref={inputRef}
        className="sr-only"
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        onChange={(event) => void submitFile(event.target.files?.[0])}
      />

      <div className="space-y-4">
        <p className="text-base text-slate-700">{state.message}</p>
        <button
          className="rounded-lg bg-blue-700 px-5 py-3 font-semibold text-white disabled:opacity-60"
          disabled={state.status === "uploading"}
          type="button"
          onClick={() => inputRef.current?.click()}
        >
          Choose PowerPoint
        </button>
      </div>
    </section>
  );
}
