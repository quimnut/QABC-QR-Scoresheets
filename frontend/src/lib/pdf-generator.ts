import { PDFDocument, PDFTextField, rgb } from "pdf-lib";
import type { Entry, StyleSet } from "./types";
import { getEntryType } from "./style-sets";
import { makeQRBytes, recommendTextSize } from "./qr";
import { padJudgingNumber } from "./judging-number";
import { initPdfWorker } from "./pdfjs-init";

// ─── Constants (converted from PyMuPDF top-left origin to pdf-lib bottom-left)
// A4 page height: 841.89 pt
// pdf-lib y = 841.89 - pymupdf_y_bottom
// ──────────────────────────────────────────────────────────────────────────

// QR code: PyMuPDF Rect(205, 720, 275, 790)
const QR_RECT = { x: 205, y: 51.89, width: 70, height: 70 } as const;

// Logo: PyMuPDF Rect(529, 6, 589, 66)
const LOGO_RECT = { x: 529, y: 775.89, width: 60, height: 60 } as const;

// Template URLs — served from public/templates/
export const TEMPLATE_URLS = {
  beer: "/templates/BSTR-1.0-QABC.pdf",
  cider: "/templates/CSTR-1.0-QABC.pdf",
  mead: "/templates/MSTR-1.0-QABC.pdf",
} as const;

// ─── Template bytes cache ─────────────────────────────────────────────────

export type TemplateBytes = {
  beer: ArrayBuffer;
  cider: ArrayBuffer;
  mead: ArrayBuffer;
};

/** Fetch all three PDF template files. Call once and reuse the result. */
export async function fetchTemplateBytes(): Promise<TemplateBytes> {
  const [beer, cider, mead] = await Promise.all([
    fetch(TEMPLATE_URLS.beer).then((r) => r.arrayBuffer()),
    fetch(TEMPLATE_URLS.cider).then((r) => r.arrayBuffer()),
    fetch(TEMPLATE_URLS.mead).then((r) => r.arrayBuffer()),
  ]);
  return { beer, cider, mead };
}

// ─── Single page generation ───────────────────────────────────────────────

/**
 * Fill one template page for a single entry.
 * Returns the flattened single-page PDF as a Uint8Array.
 */
export async function fillScoresheetPage(
  templateBytes: ArrayBuffer,
  entry: Entry,
  logoBytes: Uint8Array | null
): Promise<Uint8Array> {
  // Load a fresh copy of the template so fields start blank
  const doc = await PDFDocument.load(templateBytes);
  const form = doc.getForm();
  const page = doc.getPages()[0];

  // ── QR code ──────────────────────────────────────────────────────────
  const qrBytes = await makeQRBytes(padJudgingNumber(entry["Judging Number"]));
  const qrImage = await doc.embedPng(qrBytes);
  page.drawImage(qrImage, QR_RECT);

  // ── Competition logo (optional) ───────────────────────────────────────
  if (logoBytes) {
    page.drawRectangle({
      x: LOGO_RECT.x,
      y: LOGO_RECT.y,
      width: LOGO_RECT.width,
      height: LOGO_RECT.height,
      color: rgb(1, 1, 1),
    });
    const logoImage = await doc.embedPng(logoBytes);
    page.drawImage(logoImage, LOGO_RECT);
  }

  // ── Synthesise footer text ────────────────────────────────────────────
  const filledEntry: Entry = {
    ...entry,
    FooterText: `Location : ${entry.Location} | Table : ${entry.Table} | Flight : ${entry.Flight}`,
  };

  // ── Fill AcroForm fields ──────────────────────────────────────────────
  const specialFontSize = recommendTextSize(filledEntry.SpecialIngredients ?? "");

  for (const field of form.getFields()) {
    const name = field.getName();
    const value = filledEntry[name];
    if (value === undefined || value === null) continue;
    if (field instanceof PDFTextField) {
      const tf = form.getTextField(name);
      tf.setText(String(value));
      if (name === "SpecialIngredients") {
        tf.setFontSize(specialFontSize);
      }
    }
  }

  // Flatten: bake field values into page content (non-editable)
  form.flatten();

  return doc.save();
}

// ─── Category PDF generation ──────────────────────────────────────────────

/**
 * Generate a single output PDF for one category.
 * Entries are pre-sorted. Each entry is rendered `copies` times.
 *
 * @param entries   Entries for this category (already sorted)
 * @param templates Template bytes by type
 * @param copies    Number of copies per entry
 * @param logoBytes Optional competition logo PNG bytes
 * @param styleSet  Style set used to determine entry type (beer/mead/cider)
 * @param onProgress Progress callback (entriesProcessed, totalEntries)
 */
export async function generateCategoryPdf(
  entries: Entry[],
  templates: TemplateBytes,
  copies: number,
  logoBytes: Uint8Array | null,
  styleSet: StyleSet,
  onProgress?: (done: number, total: number) => void
): Promise<Uint8Array> {
  const outputDoc = await PDFDocument.create();
  let done = 0;

  for (const entry of entries) {
    const entryType = getEntryType(entry, styleSet);
    const templateBuf = templates[entryType];

    const filledBytes = await fillScoresheetPage(templateBuf, entry, logoBytes);

    // Copy the single filled+flattened page into the output document
    const filledDoc = await PDFDocument.load(filledBytes);
    for (let c = 0; c < copies; c++) {
      const [page] = await outputDoc.copyPages(filledDoc, [0]);
      outputDoc.addPage(page);
    }

    done++;
    onProgress?.(done, entries.length);
  }

  return outputDoc.save();
}

// ─── Sort helpers ─────────────────────────────────────────────────────────

/** Sort entries within a category (matches Python multi-key stable sort). */
export function sortCategoryEntries(entries: Entry[]): Entry[] {
  return [...entries].sort((a, b) => {
    const loc = (a.Location ?? "").localeCompare(b.Location ?? "");
    if (loc !== 0) return loc;
    const round = (a.Round ?? "").localeCompare(b.Round ?? "");
    if (round !== 0) return round;
    const table = (a.Table ?? "").localeCompare(b.Table ?? "");
    if (table !== 0) return table;
    return (a.Flight ?? "").localeCompare(b.Flight ?? "");
  });
}

// ─── Logo processing ──────────────────────────────────────────────────────

/**
 * Process an uploaded logo image: resize to 200×200 square with white
 * background, return as PNG Uint8Array.
 *
 * Runs on the main thread (requires Canvas API).
 */
export async function processLogoImage(file: File): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const TARGET = 200;
      const scale = TARGET / Math.max(img.width, img.height);
      const newW = Math.round(img.width * scale);
      const newH = Math.round(img.height * scale);

      const canvas = document.createElement("canvas");
      canvas.width = TARGET;
      canvas.height = TARGET;
      const ctx = canvas.getContext("2d")!;
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, TARGET, TARGET);
      const offsetX = Math.floor((TARGET - newW) / 2);
      const offsetY = Math.floor((TARGET - newH) / 2);
      ctx.drawImage(img, offsetX, offsetY, newW, newH);
      URL.revokeObjectURL(url);

      canvas.toBlob(
        (blob) => {
          if (!blob) {
            reject(new Error("Failed to create logo PNG blob"));
            return;
          }
          blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
        },
        "image/png"
      );
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Failed to load logo image"));
    };
    img.src = url;
  });
}

// ─── Preview: render template page to a PNG data URL ─────────────────────

/**
 * Render the first page of a PDF (by URL or bytes) to a PNG data URL.
 * Uses PDF.js under the hood. Returns a data URL suitable for <img src>.
 */
export async function renderPdfPageToDataUrl(
  pdfSource: string | Uint8Array,
  pageIndex = 0,
  scale = 1.5
): Promise<string> {
  const { getDocument, GlobalWorkerOptions } = await import("pdfjs-dist");
  await initPdfWorker(GlobalWorkerOptions);

  const loadingTask =
    typeof pdfSource === "string"
      ? getDocument({ url: pdfSource })
      : getDocument({ data: pdfSource });

  const pdfDoc = await loadingTask.promise;
  const page = await pdfDoc.getPage(pageIndex + 1); // pdf.js is 1-indexed
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement("canvas");
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  const ctx = canvas.getContext("2d")!;

  await page.render({ canvas, canvasContext: ctx, viewport }).promise;
  loadingTask.destroy();

  return canvas.toDataURL("image/png");
}
