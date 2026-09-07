// ─── pdfjs worker initialisation ─────────────────────────────────────────
//
// pdf.worker.min.mjs is NOT a true ES module (no import/export), so Chrome
// rejects it from `import()`.  We fetch the file, wrap it in a Blob with a
// known-good MIME type, and hand pdfjs a blob: URL it can always load.
//
// NOTE: pdfjs-dist's display API must run on the main thread — it spawns
// its own nested worker and falls back to a DOM-based fake worker, which
// crashes inside Web Workers (no `document`).

let _workerBlobUrl: string | null = null;

export async function initPdfWorker(
  GlobalWorkerOptions: { workerSrc: string }
): Promise<void> {
  if (!_workerBlobUrl) {
    const res = await fetch("/pdf.worker.min.mjs");
    const text = await res.text();
    const blob = new Blob([text], { type: "text/javascript" });
    _workerBlobUrl = URL.createObjectURL(blob);
  }
  GlobalWorkerOptions.workerSrc = _workerBlobUrl;
}
