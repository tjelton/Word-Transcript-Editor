import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { exportTranscriptDocx, matchHeader, parseTranscriptDocx } from '../src/docx';

const SAMPLE = new Uint8Array(readFileSync(resolve(__dirname, '../Transcript.docx')));

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Build a minimal .docx whose body contains the given paragraph XML. */
async function makeDocx(paragraphsXml: string, extra: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}"><w:body>${paragraphsXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`,
  );
  for (const [path, content] of Object.entries(extra)) zip.file(path, content);
  return zip.generateAsync({ type: 'uint8array' });
}

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

async function readPart(docx: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(docx);
  return zip.file(path)!.async('string');
}

describe('matchHeader', () => {
  it('recognises Word "timestamp speaker" lines', () => {
    expect(matchHeader('00:00:03 Speaker 1')).toMatchObject({ stamp: '00:00:03', speaker: 'Speaker 1' });
    expect(matchHeader('01:02:03 Dr Jane Smith')).toMatchObject({ stamp: '01:02:03', speaker: 'Dr Jane Smith' });
  });

  it('recognises "speaker timestamp" lines', () => {
    expect(matchHeader('Speaker 2   0:08')).toMatchObject({
      stamp: '0:08',
      speaker: 'Speaker 2',
      format: { order: 'speakerFirst', separator: '   ' },
    });
  });

  it('does not mistake spoken text for a header', () => {
    expect(matchHeader('10:30 is when we meet.')).toBeNull();
    expect(matchHeader("Let's meet at 10:30.")).toBeNull();
    expect(matchHeader('Cleanliness is the key to a healthy, happy life.')).toBeNull();
    expect(matchHeader('')).toBeNull();
  });

  it('respects a locked-in order', () => {
    expect(matchHeader('Speaker 1 0:03', 'stampFirst')).toBeNull();
  });
});

describe('parseTranscriptDocx (sample Word for the web transcript)', () => {
  it('reads segments, speakers and text', async () => {
    const parsed = await parseTranscriptDocx(SAMPLE);
    expect(parsed.hasStoredTranscription).toBe(true);
    expect(parsed.format).toEqual({ order: 'stampFirst', separator: ' ' });
    expect(parsed.segments.map(({ stamp, speaker, text }) => ({ stamp, speaker, text }))).toEqual([
      { stamp: '00:00:03', speaker: 'Speaker 1', text: 'Cleanliness is the key to a healthy, happy life.' },
      { stamp: '00:00:06', speaker: 'Speaker 1', text: 'Cleanliness.' },
      { stamp: '00:00:08', speaker: 'Speaker 2', text: 'This is a test.' },
    ]);
  });

  it('uses the precise timings from Word\'s stored transcription', async () => {
    const { segments } = await parseTranscriptDocx(SAMPLE);
    expect(segments.map((s) => s.start)).toEqual([3.23, 6.83, 8.11]);
    expect(segments.map((s) => s.end)).toEqual([6.27, 7.55, 9.309999999999999]);
  });
});

describe('exportTranscriptDocx', () => {
  it('round-trips edits to text and speakers', async () => {
    const { segments } = await parseTranscriptDocx(SAMPLE);
    segments[1].text = 'Cleanliness & tidiness <really>.';
    segments[2].speaker = 'Speaker 1';
    segments[0].speaker = 'Speaker 3';

    const out = await exportTranscriptDocx(SAMPLE, segments);
    const reparsed = await parseTranscriptDocx(out);

    expect(reparsed.segments.map(({ stamp, speaker, text, start }) => ({ stamp, speaker, text, start }))).toEqual([
      { stamp: '00:00:03', speaker: 'Speaker 3', text: 'Cleanliness is the key to a healthy, happy life.', start: 3.23 },
      { stamp: '00:00:06', speaker: 'Speaker 1', text: 'Cleanliness & tidiness <really>.', start: 6.83 },
      { stamp: '00:00:08', speaker: 'Speaker 1', text: 'This is a test.', start: 8.11 },
    ]);
  });

  it('keeps the heading, styles and section properties of the original', async () => {
    const { segments } = await parseTranscriptDocx(SAMPLE);
    const xml = await readPart(await exportTranscriptDocx(SAMPLE, segments), 'word/document2.xml');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);

    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, 'p'));
    const texts = paragraphs.map((p) => p.textContent);
    expect(texts).toEqual([
      'Transcript',
      '00:00:03 Speaker 1',
      'Cleanliness is the key to a healthy, happy life.',
      '00:00:06 Speaker 1',
      'Cleanliness.',
      '00:00:08 Speaker 2',
      'This is a test.',
    ]);
    // Heading style kept; transcript paragraphs keep Word's "Normal" style.
    expect(paragraphs[0].getElementsByTagNameNS(W_NS, 'pStyle')[0].getAttributeNS(W_NS, 'val')).toBe('Heading1');
    expect(paragraphs[1].getElementsByTagNameNS(W_NS, 'pStyle')[0].getAttributeNS(W_NS, 'val')).toBe('Normal');
    // Section properties remain the last child of the body.
    const body = doc.getElementsByTagNameNS(W_NS, 'body')[0];
    expect(body.lastElementChild?.localName).toBe('sectPr');
  });

  it('updates Word\'s stored transcription and keeps the OneDrive audio link', async () => {
    const { segments } = await parseTranscriptDocx(SAMPLE);
    segments[0].text = 'Edited text';
    segments[2].speaker = 'Speaker 1';
    segments[1].speaker = 'Interviewer';

    const xml = await readPart(await exportTranscriptDocx(SAMPLE, segments), 'customXML/item.xml');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(doc.documentElement.localName).toBe('storedTranscription');
    const json = JSON.parse(doc.documentElement.textContent!);

    expect(json.storageType).toBe('DocumentXmlStorage');
    expect(json.descriptor.audioOneDriveItem.itemId).toBe('01JQ46J6MDER6WQSNXLREJADWMDKFVURFI');
    const tr = json.descriptor.transcription;
    expect(tr.transcriptSegments).toEqual([
      { text: 'Edited text', language: 'en', start: 3.23, end: 6.27, speakerId: 0 },
      { text: 'Cleanliness.', language: 'en', start: 6.83, end: 7.55, speakerId: 2 },
      { text: 'This is a test.', language: 'en', start: 8.11, end: 9.309999999999999, speakerId: 0 },
    ]);
    expect(tr.speakerNames).toEqual([null, null, 'Interviewer']);
  });

  it('writes multi-line text as separate paragraphs and preserves whitespace', async () => {
    const { segments } = await parseTranscriptDocx(SAMPLE);
    segments[0].text = 'First line\n  Second line with leading spaces';
    const out = await exportTranscriptDocx(SAMPLE, segments);
    const xml = await readPart(out, 'word/document2.xml');
    expect(xml).toContain('<w:t xml:space="preserve">  Second line with leading spaces</w:t>');
    const reparsed = await parseTranscriptDocx(out);
    expect(reparsed.segments[0].text).toBe('First line\n  Second line with leading spaces');
  });

  it('is stable when exported repeatedly', async () => {
    const first = await exportTranscriptDocx(SAMPLE, (await parseTranscriptDocx(SAMPLE)).segments);
    const second = await exportTranscriptDocx(first, (await parseTranscriptDocx(first)).segments);
    expect(await readPart(second, 'word/document2.xml')).toBe(await readPart(first, 'word/document2.xml'));
    expect(await readPart(second, 'customXML/item.xml')).toBe(await readPart(first, 'customXML/item.xml'));
  });
});

describe('other document layouts', () => {
  it('handles documents without stored transcription data', async () => {
    const docx = await makeDocx(
      para('Meeting notes') +
        para('00:00:01 Speaker 1') +
        para('Hello there.') +
        para('00:01:05 Speaker 2') +
        para('Hi.') +
        para('') +
        para('How are you?'),
    );
    const parsed = await parseTranscriptDocx(docx);
    expect(parsed.hasStoredTranscription).toBe(false);
    expect(parsed.segments.map(({ start, speaker, text }) => ({ start, speaker, text }))).toEqual([
      { start: 1, speaker: 'Speaker 1', text: 'Hello there.' },
      { start: 65, speaker: 'Speaker 2', text: 'Hi.\n\nHow are you?' },
    ]);

    parsed.segments[0].speaker = 'Speaker 2';
    const out = await exportTranscriptDocx(docx, parsed.segments);
    const reparsed = await parseTranscriptDocx(out);
    expect(reparsed.segments.map((s) => s.speaker)).toEqual(['Speaker 2', 'Speaker 2']);
    expect(reparsed.segments[1].text).toBe('Hi.\n\nHow are you?');
    expect(await readPart(out, 'word/document.xml')).toContain('Meeting notes');
  });

  it('handles "speaker timestamp" headers with text in the same paragraph', async () => {
    const docx = await makeDocx(
      '<w:p><w:r><w:t>Speaker 1   0:03</w:t></w:r><w:r><w:br/><w:t>Good morning.</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Speaker 2   1:15</w:t></w:r><w:r><w:br/><w:t>Morning</w:t><w:tab/><w:t>all.</w:t></w:r></w:p>',
    );
    const parsed = await parseTranscriptDocx(docx);
    expect(parsed.format).toEqual({ order: 'speakerFirst', separator: '   ' });
    expect(parsed.segments.map(({ stamp, start, speaker, text }) => ({ stamp, start, speaker, text }))).toEqual([
      { stamp: '0:03', start: 3, speaker: 'Speaker 1', text: 'Good morning.' },
      { stamp: '1:15', start: 75, speaker: 'Speaker 2', text: 'Morning\tall.' },
    ]);

    const reparsed = await parseTranscriptDocx(await exportTranscriptDocx(docx, parsed.segments));
    expect(reparsed.segments.map(({ stamp, speaker, text }) => ({ stamp, speaker, text }))).toEqual([
      { stamp: '0:03', speaker: 'Speaker 1', text: 'Good morning.' },
      { stamp: '1:15', speaker: 'Speaker 2', text: 'Morning\tall.' },
    ]);
  });

  it('matches stored timings by timestamp when segment counts differ', async () => {
    const stored = {
      storageType: 'DocumentXmlStorage',
      descriptor: {
        transcription: {
          transcriptSegments: [
            { text: 'a', language: 'en', start: 1.4, end: 2, speakerId: 0 },
            { text: 'b', language: 'en', start: 5.7, end: 6, speakerId: 1 },
            { text: 'c', language: 'en', start: 9.2, end: 10, speakerId: 0 },
          ],
          speakerNames: [null, null],
        },
      },
    };
    const docx = await makeDocx(para('00:00:01 Speaker 1') + para('a') + para('00:00:09 Speaker 1') + para('c'), {
      'customXml/item1.xml': `<storedTranscription xmlns="http://schemas.microsoft.com/office/transcription/2022">${JSON.stringify(stored)}</storedTranscription>`,
    });
    const { segments } = await parseTranscriptDocx(docx);
    expect(segments.map((s) => s.start)).toEqual([1.4, 9.2]);
  });

  it('ignores tab-stop definitions in paragraph properties', async () => {
    const docx = await makeDocx(
      '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>00:00:02 Speaker 1</w:t></w:r></w:p>' +
        para('Text.'),
    );
    const { segments } = await parseTranscriptDocx(docx);
    expect(segments[0]).toMatchObject({ speaker: 'Speaker 1', text: 'Text.' });
  });

  it('reports a helpful error for documents without a transcript', async () => {
    await expect(parseTranscriptDocx(await makeDocx(para('Just some notes.')))).rejects.toThrow(/No transcript found/);
  });

  it('reports a helpful error for files that are not .docx', async () => {
    await expect(parseTranscriptDocx(new TextEncoder().encode('not a zip'))).rejects.toThrow(/not a valid Word document/);
  });
});
