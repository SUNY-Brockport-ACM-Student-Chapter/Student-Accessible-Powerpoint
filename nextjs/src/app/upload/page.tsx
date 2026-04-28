import { UploadDropzone } from "@/components/UploadDropzone";

export default function UploadPage() {
  return (
    <main className="mx-auto flex min-h-screen max-w-3xl flex-col justify-center gap-8 px-6 py-16">
      <div className="space-y-3">
        <p className="text-sm font-semibold uppercase tracking-wide text-blue-700">
          Stage 1
        </p>
        <h1 className="text-4xl font-bold tracking-tight text-slate-950">
          Upload a PowerPoint deck
        </h1>
        <p className="text-lg leading-8 text-slate-700">
          Choose a `.pptx` file up to 50 MB. The deck is stored privately before
          the Python processing service begins parsing and describing images.
        </p>
      </div>

      <UploadDropzone />
    </main>
  );
}
