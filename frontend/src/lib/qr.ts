import QRCode from "qrcode";

// ─── QR code generation ────────────────────────────────────────────────────

/**
 * Generate a QR PNG as a Uint8Array.
 *
 * Matches the Python qr_generator:
 *   version=1 (auto-fit), error_correction=H, border=3, size=360×360
 *
 * Uses OffscreenCanvas when available (Web Workers + modern main threads)
 * because QRCode.toDataURL calls document.createElement('canvas') which
 * is unavailable inside Web Workers.
 */
export async function makeQRBytes(judgingNumber: string, size = 360): Promise<Uint8Array> {
  const opts = {
    errorCorrectionLevel: "H" as const,
    width: size,
    margin: 3,
    color: { dark: "#000000", light: "#ffffff" },
  };

  if (typeof OffscreenCanvas !== "undefined") {
    // Works in Web Workers and modern main threads
    const canvas = new OffscreenCanvas(size, size);
    await QRCode.toCanvas(canvas as unknown as HTMLCanvasElement, judgingNumber, opts);
    const blob = await canvas.convertToBlob({ type: "image/png" });
    return new Uint8Array(await blob.arrayBuffer());
  }

  // Fallback for environments that have HTMLCanvasElement but no OffscreenCanvas
  const dataUrl = await QRCode.toDataURL(judgingNumber, opts);
  return dataUrlToUint8Array(dataUrl);
}

// ─── Font size recommendation ──────────────────────────────────────────────

/**
 * Recommend a font size (pt) for the SpecialIngredients AcroForm field
 * based on text length. Matches Python recommend_text_size().
 */
export function recommendTextSize(text: string): number {
  const len = text.length;
  if (len >= 135) return 2;
  if (len >= 110) return 3;
  if (len >= 90) return 4;
  if (len >= 70) return 5;
  if (len > 50) return 6;
  return 7;
}

// ─── Utility ──────────────────────────────────────────────────────────────

function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = dataUrl.split(",")[1];
  const binaryStr = atob(base64);
  const bytes = new Uint8Array(binaryStr.length);
  for (let i = 0; i < binaryStr.length; i++) {
    bytes[i] = binaryStr.charCodeAt(i);
  }
  return bytes;
}
