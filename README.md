# PowerPoint Accessibility App

A Streamlit application that enhances PowerPoint presentations for accessibility by:

- Adding AI-generated alt text to all images
- Adding comprehensive slide notes with content analysis
- Processing all shape types (text boxes, charts, tables)
- Optimizing alt text for screen readers

## Quick Start

### Using Docker (Recommended)

1. Make sure you have Docker installed
2. Set your Google API key in `docker-compose.yml`
3. Run: `docker-compose up --build`
4. Open: http://localhost:8501

### Manual Setup

1. Install dependencies: `pip install -r requirements.txt`
2. Set your Google API key: `export GOOGLE_API_KEY=your_key_here`
3. Run: `streamlit run new-app/ppt_notes.py`

## Usage

1. Upload a PowerPoint file (.pptx)
2. Wait for processing (AI analysis of images and content)
3. Download the enhanced presentation with accessibility features

## Features

- **Image Analysis**: AI-powered descriptions for all images
- **Alt Text**: Optimized alt text under 125 characters
- **Slide Notes**: Comprehensive content analysis in notes section
- **Multi-Shape Support**: Handles text boxes, charts, tables, and images
- **Accessibility Focused**: WCAG-compliant enhancements
