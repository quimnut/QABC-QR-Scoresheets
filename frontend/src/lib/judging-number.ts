/**
 * Normalise a judging number to the canonical 6-digit zero-padded form
 * (e.g. "42" → "000042"). Used for QR payloads and output filenames so
 * generated and processed scoresheets always line up.
 */
export function padJudgingNumber(value: string): string {
  return value.trim().padStart(6, "0");
}
