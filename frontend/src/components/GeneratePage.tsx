import { useState, useEffect, useRef, useCallback } from "react";
import type { Entry, Alert, ModalContent, GeneratorWorkerOutput } from "../lib/types";
import {
  loadGenerateState,
  saveGenerateState,
  clearGenerateState,
  getStorageUsage,
} from "../lib/db";
import { parseCSV, detectAndStoreStyleSet, getEntriesPreview, getUniqueCategories, formatFileSize } from "../lib/csv-parser";
import { processLogoImage, fetchTemplateBytes, TEMPLATE_URLS } from "../lib/pdf-generator";
import { renderPdfPageToDataUrl } from "../lib/pdf-processor";
import { rasterizePdf } from "../lib/pdf-rasterizer";
import AlertMessage from "./AlertMessage";
import ProgressBar from "./ProgressBar";
import CsvPreview from "./CsvPreview";
import Modal from "./Modal";
import JSZip from "jszip";

// ─── Types ─────────────────────────────────────────────────────────────────

interface GeneratedPdf {
  name: string;
  bytes: Uint8Array;
}

// ─── Component ─────────────────────────────────────────────────────────────

export default function GeneratePage() {
  // ── State ────────────────────────────────────────────────────────────────
  const [entries, setEntries] = useState<Entry[] | null>(null);
  const [csvFilename, setCsvFilename] = useState<string | null>(null);
  const [styleSetName, setStyleSetName] = useState<string | null>(null);
  const [styleSetId, setStyleSetId] = useState<number | null>(null);
  const [styleSetWarnings, setStyleSetWarnings] = useState<string[]>([]);
  const [logoBytes, setLogoBytes] = useState<Uint8Array | null>(null);
  const [logoDataUrl, setLogoDataUrl] = useState<string | null>(null);
  const [copies, setCopies] = useState(2);
  const [rasterize, setRasterize] = useState(true);
  const [singlePdf, setSinglePdf] = useState(false);
  const [generatedPdfs, setGeneratedPdfs] = useState<GeneratedPdf[]>([]);
  const [progress, setProgress] = useState<{ current: number; total: number; message: string } | null>(null);
  const [generating, setGenerating] = useState(false);
  const [showCsvPreview, setShowCsvPreview] = useState(false);
  const [alerts, setAlerts] = useState<Alert[]>([]);
  const [modal, setModal] = useState<ModalContent | null>(null);
  const [storageUsage, setStorageUsage] = useState<number>(0);

  const workerRef = useRef<Worker | null>(null);
  const rasterizeCancelRef = useRef(false);
  const csvInputRef = useRef<HTMLInputElement>(null);
  const logoInputRef = useRef<HTMLInputElement>(null);

  // ── Alerts ───────────────────────────────────────────────────────────────
  const addAlert = useCallback(
    (kind: Alert["kind"], message: string) => {
      const id = `${Date.now()}-${Math.random()}`;
      setAlerts((prev) => [...prev, { id, kind, message }]);
    },
    []
  );

  const dismissAlert = useCallback((id: string) => {
    setAlerts((prev) => prev.filter((a) => a.id !== id));
  }, []);

  const refreshStorage = useCallback(async () => {
    setStorageUsage(await getStorageUsage());
  }, []);

  // ── Load persisted state on mount ────────────────────────────────────────
  useEffect(() => {
    loadGenerateState().then((state) => {
      if (state.entries) setEntries(state.entries);
      if (state.csvFilename) setCsvFilename(state.csvFilename);
      if (state.styleSet) {
        setStyleSetName(state.styleSet.name);
        setStyleSetId(state.styleSetId);
        setStyleSetWarnings(state.styleSetWarnings ?? []);
      }
      if (state.logoBytes) {
        setLogoBytes(state.logoBytes);
        setLogoDataUrl(
          URL.createObjectURL(new Blob([new Uint8Array(state.logoBytes)], { type: "image/png" }))
        );
      }
      if (state.generatedPdfs?.length) setGeneratedPdfs(state.generatedPdfs);
      refreshStorage();
    });
  }, [refreshStorage]);

  // ── CSV upload ────────────────────────────────────────────────────────────
  async function handleCsvChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const buf = await file.arrayBuffer();
      const parsed = parseCSV(buf);
      if (parsed.length === 0) {
        addAlert("error", "CSV contains no entries.");
        return;
      }
      const detection = detectAndStoreStyleSet(parsed);
      const cats = getUniqueCategories(parsed);

      setEntries(parsed);
      setCsvFilename(file.name);
      setStyleSetId(detection.styleSet.id);
      setStyleSetName(detection.styleSet.name);
      setStyleSetWarnings(detection.warnings);

      await saveGenerateState({
        csvFilename: file.name,
        csvBytes: new Uint8Array(buf),
        entries: parsed,
        styleSet: detection.styleSet,
        styleSetId: detection.styleSet.id,
        styleSetConfidence: detection.confidence,
        styleSetWarnings: detection.warnings,
      });

      await refreshStorage();
      addAlert(
        "success",
        `Loaded ${parsed.length} entries across ${cats.length} categories. Style set: ${detection.styleSet.name}.`
      );

      if (detection.warnings.length > 0) {
        detection.warnings.forEach((w) => addAlert("warning", w));
      }
    } catch (err) {
      addAlert("error", `Failed to parse CSV: ${err}`);
    }
  }

  async function handleRemoveCsv() {
    setEntries(null);
    setCsvFilename(null);
    setStyleSetName(null);
    setStyleSetId(null);
    setStyleSetWarnings([]);
    setShowCsvPreview(false);
    await saveGenerateState({ csvFilename: null, csvBytes: null, entries: null, styleSet: null, styleSetId: null, styleSetConfidence: null, styleSetWarnings: [] });
    if (csvInputRef.current) csvInputRef.current.value = "";
  }

  // ── Logo upload ───────────────────────────────────────────────────────────
  async function handleLogoChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const processed = await processLogoImage(file);
      if (logoDataUrl) URL.revokeObjectURL(logoDataUrl);
      const url = URL.createObjectURL(new Blob([new Uint8Array(processed)], { type: "image/png" }));
      setLogoBytes(processed);
      setLogoDataUrl(url);
      await saveGenerateState({ logoBytes: processed });
      await refreshStorage();
      addAlert("success", "Logo uploaded and processed.");
    } catch (err) {
      addAlert("error", `Failed to process logo: ${err}`);
    }
  }

  async function handleRemoveLogo() {
    if (logoDataUrl) URL.revokeObjectURL(logoDataUrl);
    setLogoBytes(null);
    setLogoDataUrl(null);
    await saveGenerateState({ logoBytes: null });
    if (logoInputRef.current) logoInputRef.current.value = "";
  }

  // ── Preview blank template ────────────────────────────────────────────────
  async function handlePreviewTemplate() {
    setModal({ kind: "loading", title: "Loading template preview…" });
    try {
      const dataUrl = await renderPdfPageToDataUrl(
        new Uint8Array(await fetch(TEMPLATE_URLS.beer).then((r) => r.arrayBuffer()) as ArrayBuffer),
        0,
        1.5
      );
      setModal({ kind: "image", src: dataUrl, title: "Beer Scoresheet Template" });
    } catch {
      setModal(null);
      addAlert("error", "Failed to render template preview.");
    }
  }

  // ── Generate ──────────────────────────────────────────────────────────────
  async function handleGenerate() {
    if (!entries || !styleSetId) return;
    if (generatedPdfs.length > 0) {
      if (!window.confirm("This will replace the previously generated scoresheets. Continue?")) return;
    }

    setGenerating(true);
    rasterizeCancelRef.current = false;
    setProgress({ current: 0, total: entries.length, message: "Fetching templates…" });

    try {
      const templates = await fetchTemplateBytes();

      // Copy logo bytes before transferring (keep state valid)
      const logoCopy = logoBytes ? new Uint8Array(logoBytes) : null;

      const worker = new Worker(
        new URL("../workers/pdf-generator.worker.ts", import.meta.url),
        { type: "module" }
      );
      workerRef.current = worker;

      const transferList: Transferable[] = [templates.beer, templates.cider, templates.mead];
      if (logoCopy) transferList.push(logoCopy.buffer);

      worker.postMessage(
        {
          type: "generate",
          templateBytes: templates,
          entries,
          logoBytes: logoCopy ? logoCopy.buffer : null,
          copies,
          styleSetId,
          singlePdf,
        },
        transferList
      );

      worker.onmessage = async (evt: MessageEvent<GeneratorWorkerOutput>) => {
        const msg = evt.data;
        if (msg.type === "progress") {
          setProgress({ current: msg.current, total: msg.total, message: msg.message });
        } else if (msg.type === "complete") {
          worker.terminate();
          workerRef.current = null;

          try {
            const pdfs: GeneratedPdf[] = msg.pdfs.map((p) => ({
              name: p.name,
              bytes: new Uint8Array(p.bytes),
            }));

            if (rasterize) {
              // Rasterise each category PDF on the main thread for flat,
              // print-friendly pages (pdfjs runs its own worker for pixels).
              const totalFiles = pdfs.length;
              for (let f = 0; f < pdfs.length; f++) {
                const pdf = pdfs[f];
                pdfs[f] = {
                  name: pdf.name,
                  bytes: await rasterizePdf(
                    pdf.bytes,
                    (done, total) => {
                      setProgress({
                        current: done,
                        total,
                        message: `Rasterising ${pdf.name} (file ${f + 1}/${totalFiles})… page ${done}/${total}`,
                      });
                    },
                    () => rasterizeCancelRef.current
                  ),
                };
              }
            }

            if (rasterizeCancelRef.current) {
              setGenerating(false);
              setProgress(null);
              addAlert("warning", "Generation stopped.");
              return;
            }

            setGeneratedPdfs(pdfs);
            await saveGenerateState({ generatedPdfs: pdfs });
            setGenerating(false);
            setProgress(null);
            await refreshStorage();
            addAlert("success", `Generated ${pdfs.length} PDF file(s) for ${entries.length} entries.`);
          } catch (err) {
            setGenerating(false);
            setProgress(null);
            if (rasterizeCancelRef.current) {
              addAlert("warning", "Generation stopped.");
            } else {
              addAlert("error", `Rasterisation failed: ${err}`);
            }
          }
        } else if (msg.type === "error") {
          setGenerating(false);
          setProgress(null);
          addAlert("error", `Generation failed: ${msg.message}`);
          worker.terminate();
          workerRef.current = null;
        }
      };

      worker.onerror = (e) => {
        setGenerating(false);
        setProgress(null);
        addAlert("error", `Worker error: ${e.message}`);
        workerRef.current = null;
      };
    } catch (err) {
      setGenerating(false);
      setProgress(null);
      addAlert("error", `Failed to start generation: ${err}`);
    }
  }

  function handleStopGeneration() {
    rasterizeCancelRef.current = true;
    workerRef.current?.terminate();
    workerRef.current = null;
    setGenerating(false);
    setProgress(null);
    addAlert("warning", "Generation stopped.");
  }

  // ── View generated PDF ────────────────────────────────────────────────────
  function handleViewPdf(pdf: GeneratedPdf) {
    setModal({ kind: "pdf", bytes: pdf.bytes, title: pdf.name });
  }

  // ── Download individual PDF ───────────────────────────────────────────────
  function handleDownloadPdf(pdf: GeneratedPdf) {
    const url = URL.createObjectURL(
      new Blob([new Uint8Array(pdf.bytes)], { type: "application/pdf" })
    );
    const a = document.createElement("a");
    a.href = url;
    a.download = pdf.name;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ── Download all (ZIP or single PDF) ──────────────────────────────────────
  async function handleDownloadAll() {
    if (generatedPdfs.length === 0) return;
    if (generatedPdfs.length === 1) {
      handleDownloadPdf(generatedPdfs[0]);
      return;
    }
    addAlert("info", "Building ZIP archive…");
    try {
      const zip = new JSZip();
      for (const pdf of generatedPdfs) {
        zip.file(pdf.name, pdf.bytes);
      }
      const blob = await zip.generateAsync({ type: "blob" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "scoresheets.zip";
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      addAlert("error", `Failed to create ZIP: ${err}`);
    }
  }

  // ── Delete generated PDF ──────────────────────────────────────────────────
  async function handleDeletePdf(name: string) {
    if (!window.confirm(`Delete ${name}?`)) return;
    const updated = generatedPdfs.filter((p) => p.name !== name);
    setGeneratedPdfs(updated);
    await saveGenerateState({ generatedPdfs: updated });
  }

  async function handleClearAll() {
    if (!window.confirm("Delete all generated scoresheets?")) return;
    setGeneratedPdfs([]);
    await saveGenerateState({ generatedPdfs: [] });
  }

  async function handleReset() {
    if (!window.confirm("Clear all session data (CSV, logo, generated PDFs)?")) return;
    if (logoDataUrl) URL.revokeObjectURL(logoDataUrl);
    setEntries(null);
    setCsvFilename(null);
    setStyleSetName(null);
    setStyleSetId(null);
    setStyleSetWarnings([]);
    setLogoBytes(null);
    setLogoDataUrl(null);
    setGeneratedPdfs([]);
    setShowCsvPreview(false);
    await clearGenerateState();
    await refreshStorage();
    if (csvInputRef.current) csvInputRef.current.value = "";
    if (logoInputRef.current) logoInputRef.current.value = "";
  }

  // ── CSV preview data ──────────────────────────────────────────────────────
  const csvPreview = entries ? getEntriesPreview(entries) : null;

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <>
      <AlertMessage alerts={alerts} onDismiss={dismissAlert} />

      {/* ── Upload Files ── */}
      <div className="card">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <h2>Upload Files</h2>
          {(entries || logoBytes || generatedPdfs.length > 0) && (
            <button className="btn btn-small btn-danger" onClick={handleReset}>
              Reset Session
            </button>
          )}
        </div>
        {(entries || logoBytes || generatedPdfs.length > 0) && storageUsage > 0 && (
          <p style={{ textAlign: "right", fontSize: "0.8rem", color: "var(--text-muted)", marginTop: 2, marginBottom: 0 }}>
            Session storage: {formatFileSize(storageUsage)}
          </p>
        )}

        {/* CSV */}
        <div className="form-group">
          <label>Entries CSV (BCOEM Export)</label>
          <div className="file-upload">
            <input
              ref={csvInputRef}
              type="file"
              accept=".csv"
              onChange={handleCsvChange}
              style={{ display: entries ? "none" : "block" }}
            />
            {entries ? (
              <div className="file-upload-info">
                <span className="file-status success">{csvFilename}</span>
                {styleSetName && (
                  <small style={{ color: "var(--text-muted)", marginLeft: 8 }}>
                    Style set: <strong>{styleSetName}</strong>
                    {styleSetWarnings.length > 0 && " (warnings)"}
                  </small>
                )}
                &nbsp;
                <button
                  className="btn btn-small btn-secondary"
                  onClick={() => setShowCsvPreview((v) => !v)}
                >
                  {showCsvPreview ? "Hide Preview" : "Preview"}
                </button>
                <button
                  className="btn btn-small btn-danger"
                  onClick={handleRemoveCsv}
                  style={{ marginLeft: 8 }}
                >
                  Remove
                </button>
              </div>
            ) : (
              <div className="file-upload-info">
                <span className="file-status">No file selected</span>
              </div>
            )}
          </div>
          {showCsvPreview && csvPreview && (
            <CsvPreview headers={csvPreview.headers} rows={csvPreview.rows} />
          )}
        </div>

        {/* Logo */}
        <div className="form-group">
          <label>Competition Logo (optional — replaces QABC logo)</label>
          <div className="file-upload">
            <input
              ref={logoInputRef}
              type="file"
              accept="image/png,image/jpeg"
              onChange={handleLogoChange}
              style={{ display: logoBytes ? "none" : "block" }}
            />
            {logoBytes ? (
              <div className="file-upload-info" style={{ alignItems: "center" }}>
                <span className="file-status success">Logo uploaded</span>
                {logoDataUrl && (
                  <img
                    src={logoDataUrl}
                    alt="Logo preview"
                    className="logo-preview"
                    style={{ maxHeight: 60, marginLeft: 12 }}
                  />
                )}
                <button
                  className="btn btn-small btn-danger"
                  onClick={handleRemoveLogo}
                  style={{ marginLeft: 12 }}
                >
                  Remove
                </button>
              </div>
            ) : (
              <div className="file-upload-info">
                <span className="file-status">No logo (will use default QABC logo)</span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* ── Generate ── */}
      <div className="card">
        <h2>Generate Scoresheets</h2>

        <div className="form-row">
          <div className="form-group">
            <label htmlFor="copies-select">Copies per entry</label>
            <select
              id="copies-select"
              value={copies}
              onChange={(e) => setCopies(Number(e.target.value))}
            >
              {Array.from({ length: 10 }, (_, i) => i + 1).map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
          </div>

          <div
            className="form-group"
            style={{ display: "flex", flexDirection: "column", gap: 8, justifyContent: "flex-end", paddingBottom: 10 }}
          >
            <label className="checkbox-label">
              <input
                type="checkbox"
                checked={rasterize}
                onChange={(e) => setRasterize(e.target.checked)}
                disabled={generating}
              />
              Rasterise pages (300 DPI, flat)
            </label>
            <label className="checkbox-label">
              <input
                type="checkbox"
                checked={singlePdf}
                onChange={(e) => setSinglePdf(e.target.checked)}
                disabled={generating}
              />
              Single PDF (all categories)
            </label>
          </div>

          <div className="form-group" style={{ display: "flex", gap: 12, alignItems: "flex-end" }}>
            <button
              className="btn btn-secondary"
              onClick={handlePreviewTemplate}
              disabled={generating}
            >
              Preview Template
            </button>

            <button
              className="btn btn-primary"
              onClick={handleGenerate}
              disabled={!entries || !styleSetId || generating}
            >
              {generating ? "Generating…" : "Generate All PDFs"}
            </button>
          </div>
        </div>

        {entries && (
          <p style={{ marginTop: 8, fontSize: "0.875rem", color: "var(--text-muted)" }}>
            {entries.length} entries across {getUniqueCategories(entries).length} categories
            {styleSetName && ` · ${styleSetName}`}
          </p>
        )}

        {generating && progress && (
          <ProgressBar
            current={progress.current}
            total={progress.total}
            message={progress.message}
            onStop={handleStopGeneration}
          />
        )}
      </div>

      {/* ── Generated Files ── */}
      {generatedPdfs.length > 0 && (
        <div className="card">
          <h2>Generated Scoresheets</h2>
          <div className="file-list">
            {generatedPdfs.map((pdf) => (
              <div key={pdf.name} className="file-item">
                <span className="file-name">{pdf.name}</span>
                <span className="file-pages">{formatFileSize(pdf.bytes.length)}</span>
                <div className="file-actions">
                  <button
                    className="btn-icon"
                    title="Preview"
                    onClick={() => handleViewPdf(pdf)}
                  >
                    &#128269;
                  </button>
                  <button
                    className="btn-icon download"
                    title="Download"
                    onClick={() => handleDownloadPdf(pdf)}
                  >
                    &#11015;
                  </button>
                  <button
                    className="btn-icon danger"
                    title="Delete"
                    onClick={() => handleDeletePdf(pdf.name)}
                  >
                    &times;
                  </button>
                </div>
              </div>
            ))}
          </div>
          <div className="form-group" style={{ display: "flex", gap: 12, marginTop: 16 }}>
            <button className="btn btn-primary" onClick={handleDownloadAll}>
              {generatedPdfs.length === 1 ? "Download PDF" : "Download All (ZIP)"}
            </button>
            <button className="btn btn-danger" onClick={handleClearAll}>
              Clear All
            </button>
          </div>
        </div>
      )}

      <Modal content={modal} onClose={() => setModal(null)} />
    </>
  );
}
