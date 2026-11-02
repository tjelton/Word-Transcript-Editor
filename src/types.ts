export interface Segment {
  /** Stable id for the UI. */
  id: string;
  /** Timestamp exactly as it appears in the document (e.g. "00:00:03"). Read-only. */
  stamp: string;
  /** Start time in seconds (precise value from Word's stored transcription when available). */
  start: number;
  /** End time in seconds, when known. */
  end: number | null;
  /** Speaker label, e.g. "Speaker 1". */
  speaker: string;
  /** Spoken text. May contain "\n" for multiple paragraphs. */
  text: string;
}

/** How the speaker/timestamp line is laid out in the source document. */
export interface HeaderFormat {
  order: 'stampFirst' | 'speakerFirst';
  /** Whitespace between the timestamp and the speaker. */
  separator: string;
}

export interface ParsedTranscript {
  segments: Segment[];
  format: HeaderFormat;
  /** True when the document carries Word's hidden transcription data (precise timings). */
  hasStoredTranscription: boolean;
}
