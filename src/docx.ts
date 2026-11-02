/**
 * Reading and writing Word (.docx) transcripts.
 *
 * Word for the web inserts a transcript into the document as pairs of paragraphs:
 *
 *     00:00:03 Speaker 1
 *     Cleanliness is the key to a healthy, happy life.
 *
 * It also stores its own copy of the transcription (with precise start/end times and
 * speaker ids) as JSON in a custom XML part (customXML/item*.xml, root <storedTranscription>).
 *
 * The visible paragraphs are treated as the source of truth for text and speakers (they
 * reflect any edits already made in Word); the stored transcription supplies precise timings.
 * On export, the transcript paragraphs are regenerated and the stored transcription is updated
 * to match, so Word's transcript pane stays consistent with the document. Everything else in
 * the package (styles, headings before the transcript, the OneDrive audio link, ...) is kept.
 */
import JSZip from 'jszip';
import { STAMP_PATTERN, parseStamp } from './time';
import type { HeaderFormat, ParsedTranscript, Segment } from './types';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

type Bytes = ArrayBuffer | Uint8Array;

// ---------------------------------------------------------------------------------------------
// Header line detection
// ---------------------------------------------------------------------------------------------

const HEADER_STAMP_FIRST = new RegExp(`^\\s*(${STAMP_PATTERN})(\\s+)(\\S.{0,59}?)\\s*$`);
const HEADER_SPEAKER_FIRST = new RegExp(`^\\s*(\\S.{0,59}?)(\\s+)(${STAMP_PATTERN})\\s*$`);

interface HeaderMatch {
  stamp: string;
  speaker: string;
  format: HeaderFormat;
}

/**
 * Recognise a speaker/timestamp line such as "00:00:03 Speaker 1" or "Speaker 1  0:03".
 * Word punctuates transcribed speech, so a line whose "speaker" part ends in punctuation
 * (e.g. "10:30 is when we meet.") is treated as text, not a header.
 */
export function matchHeader(line: string, order?: HeaderFormat['order']): HeaderMatch | null {
  if (order !== 'speakerFirst') {
    const m = HEADER_STAMP_FIRST.exec(line);
    if (m && isPlausibleSpeaker(m[3])) {
      return { stamp: m[1], speaker: m[3], format: { order: 'stampFirst', separator: m[2] } };
    }
  }
  if (order !== 'stampFirst') {
    const m = HEADER_SPEAKER_FIRST.exec(line);
    if (m && isPlausibleSpeaker(m[1])) {
      return { stamp: m[3], speaker: m[1], format: { order: 'speakerFirst', separator: m[2] } };
    }
  }
  return null;
}

function isPlausibleSpeaker(s: string): boolean {
  return !/[.,!?;:…"”]$/.test(s) && !new RegExp(STAMP_PATTERN).test(s);
}

export function formatHeader(format: HeaderFormat, stamp: string, speaker: string): string {
  return format.order === 'stampFirst'
    ? `${stamp}${format.separator}${speaker}`
    : `${speaker}${format.separator}${stamp}`;
}

// ---------------------------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------------------------

function stripBom(s: string): string {
  return s.charCodeAt(0) === 0xfeff ? s.slice(1) : s;
}

function parseXml(text: string, what: string): Document {
  const doc = new DOMParser().parseFromString(stripBom(text), 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`Could not read ${what}: the XML is malformed.`);
  }
  return doc;
}

function serializeXml(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc);
  return xml.startsWith('<?xml')
    ? xml
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}

function isW(node: Node | null, localName: string): node is Element {
  return (
    !!node &&
    node.nodeType === 1 &&
    (node as Element).namespaceURI === W &&
    (node as Element).localName === localName
  );
}

function childW(el: Element, localName: string): Element | null {
  for (const child of Array.from(el.childNodes)) if (isW(child, localName)) return child;
  return null;
}

async function loadZip(data: Bytes): Promise<JSZip> {
  try {
    return await JSZip.loadAsync(data);
  } catch {
    throw new Error('This file is not a valid Word document (.docx).');
  }
}

/** Path of the main document part (usually word/document.xml; Word for the web uses word/document2.xml). */
async function mainDocumentPath(zip: JSZip): Promise<string> {
  const rels = zip.file('_rels/.rels');
  if (rels) {
    const doc = parseXml(await rels.async('string'), 'the package relationships');
    for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
      const target = rel.getAttribute('Target');
      if (target && (rel.getAttribute('Type') ?? '').endsWith('/officeDocument')) {
        const path = target.replace(/^\//, '');
        if (zip.file(path)) return path;
      }
    }
  }
  if (zip.file('word/document.xml')) return 'word/document.xml';
  throw new Error('This file does not look like a Word document (.docx).');
}

async function loadBody(zip: JSZip): Promise<{ path: string; doc: Document; body: Element }> {
  const path = await mainDocumentPath(zip);
  const doc = parseXml(await zip.file(path)!.async('string'), 'the document');
  const body = doc.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('The document has no body.');
  return { path, doc, body };
}

/** Visible text of a paragraph, split into lines at manual line breaks. */
function paragraphLines(p: Element): string[] {
  let text = '';
  for (const el of Array.from(p.getElementsByTagNameNS(W, '*'))) {
    // Only run content counts (w:tab inside w:tabs is a tab-stop definition, not text).
    if (!isW(el.parentNode, 'r')) continue;
    switch (el.localName) {
      case 't':
        text += el.textContent ?? '';
        break;
      case 'tab':
        text += '\t';
        break;
      case 'br':
      case 'cr':
        text += '\n';
        break;
    }
  }
  return text.split('\n');
}

/** Body-level paragraphs in order (descending into content controls), each with its top-level body child. */
function* bodyParagraphs(body: Element): Generator<{ p: Element; top: Element }> {
  function* walk(el: Element, top: Element | null): Generator<{ p: Element; top: Element }> {
    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType !== 1) continue;
      const c = child as Element;
      const t = top ?? c;
      if (isW(c, 'p')) yield { p: c, top: t };
      else if (isW(c, 'sdt')) {
        const content = childW(c, 'sdtContent');
        if (content) yield* walk(content, t);
      }
    }
  }
  yield* walk(body, null);
}

// ---------------------------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------------------------

interface BodyScan {
  segments: Segment[];
  format: HeaderFormat;
  /** Top-level body element containing the first transcript header. */
  firstNode: Element | null;
  headerParagraph: Element | null;
  textParagraph: Element | null;
}

let idCounter = 0;
export function newSegmentId(): string {
  idCounter += 1;
  return `seg-${Date.now().toString(36)}-${idCounter}`;
}

function scanBody(body: Element): BodyScan {
  const segments: Segment[] = [];
  const texts: string[][] = [];
  let format: HeaderFormat | null = null;
  let firstNode: Element | null = null;
  let headerParagraph: Element | null = null;
  let textParagraph: Element | null = null;

  for (const { p, top } of bodyParagraphs(body)) {
    for (const line of paragraphLines(p)) {
      const header = matchHeader(line, format?.order);
      if (header) {
        if (!format) {
          format = header.format;
          firstNode = top;
          headerParagraph = p;
        }
        segments.push({
          id: newSegmentId(),
          stamp: header.stamp,
          start: parseStamp(header.stamp),
          end: null,
          speaker: header.speaker,
          text: '',
        });
        texts.push([]);
      } else if (segments.length > 0) {
        texts[texts.length - 1].push(line);
        if (!textParagraph && line.trim() !== '') textParagraph = p;
      }
    }
  }

  segments.forEach((seg, i) => {
    const lines = texts[i];
    while (lines.length && lines[0].trim() === '') lines.shift();
    while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
    seg.text = lines.join('\n');
  });

  return {
    segments,
    format: format ?? { order: 'stampFirst', separator: ' ' },
    firstNode,
    headerParagraph,
    textParagraph,
  };
}

interface StoredSegment {
  text: string;
  start: number;
  end: number;
  speakerId: number;
  [key: string]: unknown;
}

interface StoredTranscription {
  path: string;
  doc: Document;
  json: {
    descriptor: {
      transcription: { transcriptSegments: StoredSegment[]; speakerNames?: (string | null)[] };
    };
  };
}

async function findStoredTranscription(zip: JSZip): Promise<StoredTranscription | null> {
  for (const path of Object.keys(zip.files)) {
    if (!/^customxml\/item\d*\.xml$/i.test(path)) continue;
    const xml = await zip.file(path)!.async('string');
    if (!xml.includes('storedTranscription')) continue;
    try {
      const doc = parseXml(xml, path);
      if (doc.documentElement.localName !== 'storedTranscription') continue;
      const json = JSON.parse(doc.documentElement.textContent ?? '');
      const segs = json?.descriptor?.transcription?.transcriptSegments;
      if (Array.isArray(segs)) return { path, doc, json };
    } catch {
      // Unreadable stored transcription: carry on without precise timings.
    }
  }
  return null;
}

/** Copy precise start/end times from Word's stored transcription onto the parsed segments. */
function applyStoredTimings(segments: Segment[], stored: StoredSegment[]): void {
  const close = (seg: Segment, s: StoredSegment) =>
    typeof s.start === 'number' && Math.abs(s.start - parseStamp(seg.stamp)) < 1.5;
  const copy = (seg: Segment, s: StoredSegment) => {
    seg.start = s.start;
    if (typeof s.end === 'number') seg.end = s.end;
  };

  if (stored.length === segments.length && segments.every((seg, i) => close(seg, stored[i]))) {
    segments.forEach((seg, i) => copy(seg, stored[i]));
    return;
  }
  // Counts differ (segments added/removed in Word): match greedily by timestamp.
  let j = 0;
  for (const seg of segments) {
    const t = parseStamp(seg.stamp);
    while (j < stored.length && stored[j].start < t - 1.5) j++;
    if (j < stored.length && close(seg, stored[j])) copy(seg, stored[j++]);
  }
}

export async function parseTranscriptDocx(data: Bytes): Promise<ParsedTranscript> {
  const zip = await loadZip(data);
  const { body } = await loadBody(zip);
  const scan = scanBody(body);
  if (scan.segments.length === 0) {
    throw new Error(
      'No transcript found in this document. Expected lines like "00:00:03 Speaker 1" followed by the spoken text.',
    );
  }
  const stored = await findStoredTranscription(zip);
  if (stored) applyStoredTimings(scan.segments, stored.json.descriptor.transcription.transcriptSegments);
  return {
    segments: scan.segments,
    format: scan.format,
    hasStoredTranscription: stored !== null,
  };
}

// ---------------------------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------------------------

interface ParagraphTemplate {
  pPr: Element | null;
  rPr: Element | null;
}

function templateFrom(p: Element | null): ParagraphTemplate {
  if (!p) return { pPr: null, rPr: null };
  const pPr = childW(p, 'pPr')?.cloneNode(true) as Element | undefined;
  // Never copy a section break along with the paragraph formatting.
  if (pPr) for (const s of Array.from(pPr.getElementsByTagNameNS(W, 'sectPr'))) s.remove();
  const run = p.getElementsByTagNameNS(W, 'r')[0];
  const rPr = run ? childW(run, 'rPr') : null;
  return { pPr: pPr ?? null, rPr: (rPr?.cloneNode(true) as Element) ?? null };
}

function makeParagraph(doc: Document, tpl: ParagraphTemplate, text: string): Element {
  const p = doc.createElementNS(W, 'w:p');
  if (tpl.pPr) p.appendChild(tpl.pPr.cloneNode(true));
  if (text !== '') {
    const r = doc.createElementNS(W, 'w:r');
    if (tpl.rPr) r.appendChild(tpl.rPr.cloneNode(true));
    text.split('\t').forEach((part, i) => {
      if (i > 0) r.appendChild(doc.createElementNS(W, 'w:tab'));
      if (part !== '') {
        const t = doc.createElementNS(W, 'w:t');
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        t.textContent = part;
        r.appendChild(t);
      }
    });
    p.appendChild(r);
  }
  return p;
}

/** Map a speaker label to Word's numeric speaker id, extending speakerNames as needed. */
function speakerIdResolver(names: (string | null)[]): (speaker: string) => number {
  return (speaker) => {
    const existing = names.findIndex((n, i) => (n ?? `Speaker ${i + 1}`) === speaker);
    if (existing >= 0) return existing;
    const m = /^Speaker (\d+)$/.exec(speaker);
    if (m) {
      const id = Number(m[1]) - 1;
      while (names.length <= id) names.push(null);
      if (names[id] === null) return id;
    }
    names.push(speaker);
    return names.length - 1;
  };
}

function updateStoredTranscription(stored: StoredTranscription, segments: Segment[]): string {
  const tr = stored.json.descriptor.transcription;
  const originals = new Map<number, StoredSegment>();
  for (const s of tr.transcriptSegments) originals.set(s.start, s);
  const fallbackLanguage = tr.transcriptSegments.find((s) => s.language)?.language ?? 'en';

  const names = Array.isArray(tr.speakerNames) ? [...tr.speakerNames] : [];
  const idFor = speakerIdResolver(names);

  tr.transcriptSegments = segments.map((seg, i) => {
    const original = originals.get(seg.start);
    const next = segments[i + 1];
    return {
      language: fallbackLanguage,
      ...original,
      text: seg.text,
      start: seg.start,
      end: seg.end ?? (next ? next.start : seg.start),
      speakerId: idFor(seg.speaker),
    };
  });
  tr.speakerNames = names;

  stored.doc.documentElement.textContent = JSON.stringify(stored.json);
  return serializeXml(stored.doc);
}

/**
 * Build an edited copy of the original .docx: the transcript paragraphs are replaced with
 * `segments` (keeping the original paragraph/run formatting), and Word's stored transcription
 * is updated to match. Everything before the transcript (e.g. the "Transcript" heading) is kept.
 */
export async function exportTranscriptDocx(original: Bytes, segments: Segment[]): Promise<Uint8Array> {
  const zip = await loadZip(original);
  const { path, doc, body } = await loadBody(zip);
  const scan = scanBody(body);
  if (!scan.firstNode) throw new Error('The original document no longer contains a transcript.');

  const headerTpl = templateFrom(scan.headerParagraph);
  const textTpl = scan.textParagraph ? templateFrom(scan.textParagraph) : headerTpl;

  // Remove the old transcript (everything from the first header to the section properties),
  // keeping any tables that were in that range.
  const sectPr = Array.from(body.childNodes).find((n) => isW(n, 'sectPr')) ?? null;
  const kept: Node[] = [];
  let node: Node | null = scan.firstNode;
  while (node && node !== sectPr) {
    const next: Node | null = node.nextSibling;
    if (isW(node, 'tbl')) kept.push(node);
    body.removeChild(node);
    node = next;
  }

  const fragment = doc.createDocumentFragment();
  for (const seg of segments) {
    fragment.appendChild(makeParagraph(doc, headerTpl, formatHeader(scan.format, seg.stamp, seg.speaker)));
    for (const line of seg.text.split('\n')) fragment.appendChild(makeParagraph(doc, textTpl, line));
  }
  for (const k of kept) fragment.appendChild(k);
  body.insertBefore(fragment, sectPr);

  zip.file(path, serializeXml(doc));

  const stored = await findStoredTranscription(zip);
  if (stored) zip.file(stored.path, updateStoredTranscription(stored, segments));

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: DOCX_MIME });
}

export { DOCX_MIME };

/** "Interview.docx" -> "Interview (edited).docx" (without stacking "(edited)" on re-exports). */
export function editedFileName(name: string): string {
  const base = name.replace(/\.docx$/i, '').replace(/ \(edited\)$/, '');
  return `${base} (edited).docx`;
}
