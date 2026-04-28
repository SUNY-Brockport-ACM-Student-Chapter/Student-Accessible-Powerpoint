"use client";

import { useRouter } from "next/navigation";
import { useRef, useState } from "react";

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

    setState({ status: "uploading", message: "Uploading deck..." });

    const formData = new FormData();
    formData.append("file", file);

    const response = await fetch("/api/uploads", {
      method: "POST",
      body: formData,
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
