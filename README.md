# QABC Scoresheet Manager

Web-based scoresheet management tool for BJCP-sanctioned beer, cider, and mead competitions using BCOEM.

## Features

- **Scoresheet Generation**: Upload BCOEM entries CSV, generate PDFs with QR codes for each entry
- **Custom Logo**: Optionally upload a competition logo to replace the default QABC branding
- **Scoresheet Processing**: Upload scanned scoresheets, automatically sort pages by QR code / judging number
- **Reject Handling**: Manual review interface for pages where QR code couldn't be read — assign a judging number or delete the page
- **Client-side only**: All PDF generation, QR encoding/decoding, and ZIP packaging run in the browser — no server processing at runtime

## Quick Start with Docker

```bash
docker compose up -d
```

Access at http://localhost:8080

### Docker Configuration

The app is served by nginx on port 80 inside the container. Map to any host port:

```yaml
ports:
  - "8080:80"   # Access via localhost:8080
```

For macvlan/direct network access, remove port mapping and assign IP directly.

## Local Development

### Requirements

- Node.js 22.13+ (Node 24 LTS recommended)
- npm

### Installation

```bash
cd frontend
npm install
```

### Running Locally

```bash
cd frontend
npm run dev
```

Opens at http://localhost:4321 with hot-reload enabled.

### Building

```bash
cd frontend
npm run build
```

Static output is written to `frontend/dist/`. Serve with any static file server.

### Type Checking

```bash
cd frontend
npx tsc --noEmit
```

(`astro check` is not recommended — it OOMs on large type definitions from pdfjs-dist.)

## Usage

### Scoresheet Generation

1. Navigate to the **Generate Scoresheets** tab
2. Upload your BCOEM entries CSV export (with flights assigned)
3. Optionally upload a competition logo (PNG/JPG) to replace the default QABC logo
4. Choose options: **Rasterise pages** (300 DPI flat pages for printing, on by default) and **Single PDF** (one combined `scoresheets_all.pdf` instead of per-category files)
5. Click **Generate All PDFs**
6. Download individual PDFs or all as a ZIP

### Scoresheet Processing

1. Navigate to the **Process Scoresheets** tab
2. Upload one or more scanned scoresheet PDFs (multi-page supported)
3. Click **Process Scoresheets**
4. Review any rejected pages (where QR code couldn't be decoded)
   - Enter a judging number and click **Assign**, or click **Delete Page** to discard
5. Download sorted PDFs (named by judging number, e.g. `000042.pdf`) or all as a ZIP

## PDF Templates

BJCP-compliant scoresheet templates are included in `frontend/public/templates/`:

| File | Type |
|------|------|
| `BSTR-1.0-QABC.pdf` | Beer Scoresheet |
| `CSTR-1.0-QABC.pdf` | Cider Scoresheet |
| `MSTR-1.0-QABC.pdf` | Mead Scoresheet |

### Dynamic Style Set Detection

Templates are automatically selected based on the detected competition style set:

- **Supported Style Sets:** BJCP 2008, 2015, 2021, 2025 — AABC 2022, 2025 — Brewers Association (BA)
- The system analyses category patterns in the uploaded CSV, auto-detects the style set, and assigns the correct Beer/Mead/Cider template

For unknown or ambiguous style sets the system defaults to BJCP 2021.

## Project Structure

```
QABC-QR-Scoresheets/
├── frontend/                    # Astro project (static site generator)
│   ├── public/
│   │   ├── templates/           # PDF scoresheet templates
│   │   └── pdf.worker.min.mjs   # pdfjs-dist worker (served as static asset)
│   ├── src/
│   │   ├── lib/
│   │   │   ├── csv-parser.ts    # BCOEM CSV parsing (PapaParse)
│   │   │   ├── db.ts            # IndexedDB session storage (idb)
│   │   │   ├── pdf-generator.ts # AcroForm fill + flatten (pdf-lib)
│   │   │   ├── pdf-processor.ts # QR decode + page sort (pdfjs-dist, jsqr)
│   │   │   ├── pdf-rasterizer.ts# 300 DPI page rasterisation for print (pdfjs-dist, pdf-lib)
│   │   │   ├── pdfjs-init.ts    # pdfjs worker blob: URL bootstrap
│   │   │   ├── judging-number.ts# Zero-padded judging number helper
│   │   │   ├── qr.ts            # QR code generation (qrcode, OffscreenCanvas)
│   │   │   └── style-sets.ts    # Style set definitions
│   │   ├── workers/
│   │   │   └── pdf-generator.worker.ts  # Web Worker for PDF generation
│   │   ├── components/
│   │   │   ├── GeneratePage.tsx # Generate tab UI
│   │   │   ├── ProcessPage.tsx  # Process tab UI
│   │   │   └── Modal.tsx        # Multi-page PDF viewer
│   │   ├── layouts/
│   │   │   └── Layout.astro     # Base HTML layout
│   │   └── pages/
│   │       ├── index.astro      # Redirects to /generate
│   │       ├── generate.astro   # Generate tab
│   │       └── process.astro    # Process tab
│   ├── astro.config.mjs
│   ├── tsconfig.json
│   └── package.json
├── Dockerfile                   # Two-stage: Node 24 builder → nginx 1.27-alpine
├── docker-compose.yml
└── nginx.conf                   # Serves Astro dist; .mjs MIME fix included
```

## Tech Stack

- **Framework**: [Astro](https://astro.build/) (SSG, `output: "static"`)
- **UI**: React 19 + TypeScript
- **PDF Generation**: [pdf-lib](https://pdf-lib.js.org/) — AcroForm field fill + flatten, then 300 DPI rasterisation via [pdf.js](https://mozilla.github.io/pdf.js/) for flat, print-friendly pages
- **PDF Rendering**: [pdfjs-dist](https://mozilla.github.io/pdf.js/) v6 — page rasterisation for viewer/QR decode
- **QR Generation**: [qrcode](https://www.npmjs.com/package/qrcode) via OffscreenCanvas (Web Worker safe)
- **QR Decoding**: [jsqr](https://github.com/cozmo/jsQR) — multi-strategy pipeline (crop, threshold, invert, rotate)
- **CSV Parsing**: [PapaParse](https://www.papaparse.com/)
- **ZIP Packaging**: [JSZip](https://stuk.github.io/jszip/)
- **Session Storage**: [idb](https://github.com/jakearchibald/idb) (IndexedDB wrapper) — persists across page refresh
- **Container**: nginx 1.27-alpine serving static files

## Acknowledgments

This project utilises style set data and references from the following sources:

- **[Beer Judge Certification Program (BJCP)](https://www.bjcp.org/)** — BJCP 2008, 2015, 2021, and 2025 style guidelines
- **[Australian Amateur Brewing Championship (AABC)](https://aabc.asn.au/)** — AABC 2022 and 2025 style guidelines
- **[Brew Competition Online Entry & Management (BCOEM)](https://brewingcompetitions.com/)** — Competition management platform and style set data structure
  - Repository: [github.com/geoffhumphrey/brewcompetitiononlineentry](https://github.com/geoffhumphrey/brewcompetitiononlineentry)
  - License: [GPL](https://brewingcompetitions.com/license)

We are grateful to these organisations and projects for their contributions to the homebrewing community.

## License

BJCP Scoresheet Copyright 2018 Beer Judge Certification Program.
