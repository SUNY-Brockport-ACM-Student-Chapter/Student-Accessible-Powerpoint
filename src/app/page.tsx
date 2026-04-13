"use client";
import React, { useState } from "react";
import {
  Upload,
  FileType,
  CheckCircle2,
  ChevronRight,
  AlertCircle,
  Loader2,
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";

type Stage = "upload" | "analyzing" | "review" | "download";

export default function AccessibilityApp() {
  const [stage, setStage] = useState<Stage>("upload");
  const [file, setFile] = useState<File | null>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) {
      const selectedFile = e.target.files[0];
      setFile(selectedFile);
      setStage("analyzing");

      const formData = new FormData();
      formData.append("file", selectedFile);

      try {
        const response = await fetch("/api/process", {
          method: "POST",
          body: formData,
        });

        if (response.ok) {
          // In a real flow, we'd get the parsed slides here
          setTimeout(() => setStage("review"), 2000);
        } else {
          console.error("Failed to process");
          setStage("upload");
        }
      } catch (err) {
        console.error(err);
        setStage("upload");
      }
    }
  };

  return (
    <div className="min-h-screen bg-[#050505] text-white font-sans selection:bg-purple-500/30">
      {/* Background decoration */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-purple-900/20 blur-[120px] rounded-full" />
        <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-blue-900/20 blur-[120px] rounded-full" />
      </div>

      <main className="relative z-10 max-w-5xl mx-auto px-6 py-16">
        {/* Header */}
        <header className="mb-16">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="flex items-center gap-3 mb-4"
          >
            <div className="p-2 bg-gradient-to-br from-purple-600 to-blue-600 rounded-xl">
              <Upload className="w-6 h-6" />
            </div>
            <span className="text-sm font-medium tracking-wider uppercase text-zinc-400">
              Accessibility Tool
            </span>
          </motion.div>
          <motion.h1
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className="text-5xl md:text-7xl font-bold tracking-tight mb-6 bg-gradient-to-b from-white to-zinc-500 bg-clip-text text-transparent"
          >
            Student Accessible <br />
            PowerPoints
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.2 }}
            className="text-xl text-zinc-400 max-w-2xl leading-relaxed"
          >
            Transform classroom materials into WCAG-compliant slides using
            RAG-enhanced AI. Automate alt-text, generate accessible notes, and
            ensure every student can learn.
          </motion.p>
        </header>

        {/* Stepper */}
        <div className="flex items-center gap-4 mb-12 overflow-x-auto pb-4 no-scrollbar">
          <StepItem
            active={stage === "upload"}
            completed={stage !== "upload"}
            label="Upload"
            icon={<FileType size={18} />}
          />
          <ChevronRight className="text-zinc-700 shrink-0" size={16} />
          <StepItem
            active={stage === "analyzing"}
            completed={["review", "download"].includes(stage)}
            label="AI Analysis"
            icon={
              <Loader2
                size={18}
                className={stage === "analyzing" ? "animate-spin" : ""}
              />
            }
          />
          <ChevronRight className="text-zinc-700 shrink-0" size={16} />
          <StepItem
            active={stage === "review"}
            completed={["download"].includes(stage)}
            label="Review"
            icon={<CheckCircle2 size={18} />}
          />
          <ChevronRight className="text-zinc-700 shrink-0" size={16} />
          <StepItem
            active={stage === "download"}
            completed={false}
            label="Download"
            icon={<Upload size={18} />}
          />
        </div>

        {/* Main Interface */}
        <AnimatePresence mode="wait">
          {stage === "upload" && (
            <motion.div
              key="upload"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="relative group"
            >
              <label className="block w-full cursor-pointer">
                <div className="absolute -inset-1 bg-gradient-to-r from-purple-600 to-blue-600 rounded-[2rem] blur opacity-25 group-hover:opacity-40 transition duration-500" />
                <div className="relative bg-[#0A0A0A] border border-white/10 rounded-[2rem] p-16 flex flex-col items-center justify-center text-center transition-colors hover:border-white/20">
                  <div className="w-20 h-20 bg-zinc-900 rounded-3xl flex items-center justify-center mb-6 group-hover:scale-110 transition-transform duration-500">
                    <Upload className="w-10 h-10 text-purple-500" />
                  </div>
                  <h3 className="text-2xl font-semibold mb-2 text-zinc-100">
                    Drop your presentation
                  </h3>
                  <p className="text-zinc-500 mb-8">
                    Support .pptx files up to 50MB
                  </p>
                  <div className="px-8 py-4 bg-white text-black font-bold rounded-2xl hover:bg-zinc-200 transition-colors">
                    Browse Files
                  </div>
                </div>
                <input
                  type="file"
                  className="hidden"
                  accept=".pptx"
                  onChange={handleFileUpload}
                />
              </label>
            </motion.div>
          )}

          {stage === "analyzing" && (
            <motion.div
              key="analyzing"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="bg-[#0A0A0A] border border-white/10 rounded-[2rem] p-24 flex flex-col items-center justify-center text-center"
            >
              <div className="relative w-32 h-32 mb-12">
                <div className="absolute inset-0 bg-purple-600/20 rounded-full blur-2xl animate-pulse" />
                <div className="relative w-full h-full border-4 border-zinc-800 rounded-full flex items-center justify-center">
                  <div className="absolute inset-[-4px] border-4 border-t-purple-500 rounded-full animate-spin" />
                  <Loader2 className="w-12 h-12 text-zinc-400" />
                </div>
              </div>
              <h3 className="text-2xl font-bold mb-4 tracking-tight">
                Processing Slides...
              </h3>
              <p className="text-zinc-400 max-w-md">
                Our RAG-enhanced AI is currently identifying visual elements and
                generating high-quality accessibility descriptions.
              </p>
            </motion.div>
          )}

          {stage === "review" && (
            <motion.div
              key="review"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <div className="flex justify-between items-end mb-8">
                <div>
                  <h2 className="text-3xl font-bold mb-2">
                    Review Descriptions
                  </h2>
                  <p className="text-zinc-400">
                    Verify and improve AI-generated alt-text for 12 detected
                    images.
                  </p>
                </div>
                <button
                  onClick={() => setStage("download")}
                  className="px-6 py-3 bg-purple-600 hover:bg-purple-500 text-white font-bold rounded-xl transition-colors shadow-lg shadow-purple-900/20"
                >
                  Confirm & Export
                </button>
              </div>

              {/* Mockup side-by-side review */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {[1, 2].map((i) => (
                  <div
                    key={i}
                    className="bg-[#0A0A0A] border border-white/10 rounded-[1.5rem] overflow-hidden group"
                  >
                    <div className="aspect-video bg-zinc-900/50 flex items-center justify-center relative border-b border-white/5">
                      <span className="text-zinc-600 font-mono text-sm leading-none p-2 border border-zinc-800 rounded uppercase">
                        Slide {i} Image Preview
                      </span>
                    </div>
                    <div className="p-6">
                      <div className="flex items-center gap-2 mb-4">
                        <span className="text-xs font-bold px-2 py-1 bg-purple-500/10 text-purple-400 rounded">
                          AI GENERATED
                        </span>
                        <span className="text-xs text-zinc-500">
                          Confidence: 98%
                        </span>
                      </div>
                      <textarea
                        className="w-full bg-zinc-900/50 border border-white/5 rounded-xl p-4 text-sm text-zinc-300 focus:outline-none focus:border-purple-500/50 transition-colors"
                        rows={3}
                        defaultValue="A detailed bar chart showing the growth of student enrollment from 2020 to 2024, with a significant 15% uptick in the final year."
                      />
                    </div>
                  </div>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Modern footer */}
      <footer className="mt-32 py-12 border-t border-white/5 relative z-10">
        <div className="max-w-5xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-8">
          <div className="flex items-center gap-2 text-zinc-500 text-sm">
            <AlertCircle size={16} />
            Built for SUNY Brockport ACM Student Chapter
          </div>
          <div className="flex gap-12 text-sm text-zinc-500">
            <a href="#" className="hover:text-white transition-colors">
              Documentation
            </a>
            <a href="#" className="hover:text-white transition-colors">
              Accessibility Policy
            </a>
            <a href="#" className="hover:text-white transition-colors">
              GitHub
            </a>
          </div>
        </div>
      </footer>
    </div>
  );
}

function StepItem({
  active,
  completed,
  label,
  icon,
}: {
  active: boolean;
  completed: boolean;
  label: string;
  icon: React.ReactNode;
}) {
  return (
    <div
      className={`flex items-center gap-2 px-4 py-2 rounded-full transition-all shrink-0 ${active ? "bg-white/10 text-white" : completed ? "text-purple-500" : "text-zinc-600"}`}
    >
      <div
        className={`w-8 h-8 rounded-full flex items-center justify-center ${active ? "bg-purple-600" : completed ? "bg-purple-600/20" : "bg-zinc-900"}`}
      >
        {completed ? <CheckCircle2 size={16} /> : icon}
      </div>
      <span className="text-sm font-semibold">{label}</span>
    </div>
  );
}
