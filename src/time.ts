/** Matches the timestamps Word writes into transcripts: "00:00:03", "0:03", "1:02:03" (optionally with a fraction). */
export const STAMP_PATTERN = '\\d{1,2}:\\d{2}(?::\\d{2})?(?:[.,]\\d+)?';

/** Parse a transcript timestamp ("00:01:05", "1:05", "1:05.5") into seconds. Returns NaN if it isn't one. */
export function parseStamp(stamp: string): number {
  const m = new RegExp(`^\\s*(${STAMP_PATTERN})\\s*$`).exec(stamp);
  if (!m) return NaN;
  const parts = m[1].replace(',', '.').split(':').map(Number);
  return parts.reduce((total, part) => total * 60 + part, 0);
}

/** Format seconds for the audio player: "4:05" or "1:02:03". */
export function formatClock(seconds: number): string {
  if (!Number.isFinite(seconds) || seconds < 0) seconds = 0;
  const s = Math.floor(seconds);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = String(s % 60).padStart(2, '0');
  return h > 0 ? `${h}:${String(m).padStart(2, '0')}:${sec}` : `${m}:${sec}`;
}

/**
 * Format seconds as a transcript timestamp in the same style as `example`
 * ("00:00:03" -> "00:01:05", "0:03" -> "1:05"). Hours are added when needed.
 */
export function formatStamp(seconds: number, example: string): string {
  const s = Math.max(0, Math.floor((Number.isFinite(seconds) ? seconds : 0) + 1e-6));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const pad2 = (n: number) => String(n).padStart(2, '0');
  const parts = example.trim().split(':');
  if (parts.length === 3 || h > 0) {
    const hourWidth = parts.length === 3 ? parts[0].length : 1;
    return `${String(h).padStart(hourWidth, '0')}:${pad2(m)}:${pad2(s % 60)}`;
  }
  return `${String(m).padStart(parts[0].length, '0')}:${pad2(s % 60)}`;
}

/**
 * Index of the segment playing at time `t`: the last segment whose start is <= t.
 * `starts` must be in ascending order. Returns -1 before the first segment.
 */
export function activeSegmentIndex(starts: readonly number[], t: number): number {
  let lo = 0;
  let hi = starts.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (starts[mid] <= t + 1e-6) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}
