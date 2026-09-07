import { PDFDocument } from "pdf-lib";
import { initPdfWorker } from "./pdfjs-init";

// ─── Constants ─────────────────────────────────────────────────────────────

/** Rasterisation resolution for generated pages (print standard). */
const RASTER_DPI = 300;

/** JPEG quality for rasterised pages (white sheets with text compress well). */
const RASTER_JPEG_QUALITY = 0.9;

// ─── Main-thread rasterisation ─────────────────────────────────────────────
//
// pdfjs-dist's display API must run on the main thread: it spawns its own
// pdf.worker and falls back to a DOM-based fake worker otherwise, which
// crashes inside Web Workers (no `document`).  The heavy pixel work still
// happens in pdfjs's real worker, so this keeps the UI responsive.

/**
 * Rebuild a PDF so every page is a single flat JPEG at 300 DPI.
 * Produces print-shop-friendly image-only pages.
 *
 * @param pdfBytes    Source PDF (vector pages are discarded after rasterising)
 * @param onProgress  Called after each page: (pagesDone, totalPages)
 * @param cancelCheck Called between pages; return true to abort
 */
export async function rasterizePdf(
  pdfBytes: Uint8Array,
  onProgress?: (done: number, total: number) => void | Promise<void>,
  cancelCheck?: () => boolean
): Promise<Uint8Array> {
  const { getDocument, GlobalWorkerOptions } = await import("pdfjs-dist");
  await initPdfWorker(GlobalWorkerOptions);

  // Pass a copy: pdfjs transfers the ArrayBuffer to its worker
  const loadingTask = getDocument({ data: pdfBytes.slice() });
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;

  const outputDoc = await PDFDocument.create();

  try {
    for (let i = 0; i < numPages; i++) {
      if (cancelCheck?.()) {
        throw new Error("Rasterisation cancelled");
      }

      const page = await pdf.getPage(i + 1); // pdf.js is 1-indexed
      const base = page.getViewport({ scale: 1 });
      const viewport = page.getViewport({ scale: RASTER_DPI / 72 });

      const canvas = document.createElement("canvas");
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext("2d")!;

      await page.render({ canvas, canvasContext: ctx, viewport }).promise;

      const jpegBytes = await canvasToJpegBytes(canvas);
      const image = await outputDoc.embedJpg(jpegBytes);

      const outPage = outputDoc.addPage([base.width, base.height]);
      outPage.drawImage(image, {
        x: 0,
        y: 0,
        width: base.width,
        height: base.height,
      });

      onProgress?.(i + 1, numPages);
      // Yield to the event loop so the UI can update between pages
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    }
  } finally {
    loadingTask.destroy();
  }

  return outputDoc.save();
}

// ─── Canvas → JPEG bytes ───────────────────────────────────────────────────

function canvasToJpegBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) {
          reject(new Error("Failed to export page image"));
          return;
        }
        blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
      },
      "image/jpeg",
      RASTER_JPEG_QUALITY
    );
  });
}
