# PPTgen V2 Studio

Local PPT workflow studio for turning long-form content and reference files into slide-ready image pages, then exporting a PowerPoint deck.

## Requirements

- Windows 10/11
- Node.js 18+; Node.js 20+ is recommended
- DashScope / Qwen API key for workflow planning and text reasoning
- Google Gemini API key for Nano Banana image generation

## Start

```bash
npm install
npm start
```

Open:

```text
http://localhost:3000/
```

On Windows, you can also double-click:

```text
start-app.bat
```

The root route redirects to the active V2 UI at `/v2/index.html`.

## Active Features

- V2 smart PPT workflow UI
- Theme definition, content splitting, page preparation, single-page and batch image generation
- Reference uploads for text, Markdown, JSON, HTML, XML, DOCX, PPTX, XLS/XLSX, PDF, and images
- Image references are kept as visual context for the model; OCR is not performed
- Local generated image cache in `generated-images/`
- PPT export through the project dependency `pptxgenjs`

## Local Runtime Data

These folders are runtime artifacts and should not be committed:

- `generated-images/`
- `exports/`
- `data/studio-library.json`
- `data/reference-assets/`
- `tmp/`

Clean old local artifacts with:

```bash
npm run clean:artifacts
```
