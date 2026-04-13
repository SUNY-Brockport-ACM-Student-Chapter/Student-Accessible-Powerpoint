# PowerPoint Accessibility Enhancer (React/TypeScript)

This repository has been migrated from a Python/Streamlit implementation to a modern **Next.js 14** application with **TypeScript** and **Tailwind CSS**.

## 🚀 Transitioned Features
- **Frontend**: Full React UI using Framer Motion for premium animations and Lucide for iconography.
- **Type Safety**: Fully typed logic for PPTX parsing and AI interactions.
- **RAG Integration**: TypeScript client for the existing ChromaDB API.
- **AI Descriptions**: Integrated Google Gemini Vision for alt-text generation.
- **Next.js Architecture**: Secure API routes for server-side processing.

## 📁 New Project Structure
- `src/app/`: Next.js App Router (UI and API routes).
- `src/lib/`: Core logic (PPTX processing, Gemini service, Chroma client).
- `backups/python-legacy/`: Original Python/Streamlit implementation for reference.

## 🛠️ Getting Started

### 1. Prerequisites
- Node.js 18+
- The existing ChromaDB API running (see `backups/python-legacy/start_app.py` for instructions)

### 2. Setup
```bash
npm install
```

### 3. Environment Variables
Create a `.env.local` file in the root:
```env
GOOGLE_API_KEY=your_gemini_api_key_ here
CHROMA_API_URL=http://localhost:8001
```

### 4. Running the App
```bash
npm run dev
```

## 📖 How it Works
1. **Upload**: Users drop a `.pptx` file into the modern React dashboard.
2. **Analysis**: The Next.js API route parses the presentation using `adm-zip` and calls Gemini Vision for image analysis.
3. **Review**: Users can verify and edit the AI-generated alt-text in a sleek side-by-side interface.
4. **Export**: The updated presentation (with native alt-text and enhanced slide notes) is generated and downloaded.
