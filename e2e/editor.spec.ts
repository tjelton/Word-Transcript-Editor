import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';

const DOCX = resolve(import.meta.dirname, '../Transcript.docx');

/** A 20-second mono 8 kHz sine-wave WAV, generated so no binary fixture is needed. */
function makeWav(seconds = 20, rate = 8000): Buffer {
  const samples = seconds * rate;
  const buf = Buffer.alloc(44 + samples * 2);
  buf.write('RIFF', 0);
  buf.writeUInt32LE(36 + samples * 2, 4);
  buf.write('WAVEfmt ', 8);
  buf.writeUInt32LE(16, 16);
  buf.writeUInt16LE(1, 20); // PCM
  buf.writeUInt16LE(1, 22); // mono
  buf.writeUInt32LE(rate, 24);
  buf.writeUInt32LE(rate * 2, 28);
  buf.writeUInt16LE(2, 32);
  buf.writeUInt16LE(16, 34);
  buf.write('data', 36);
  buf.writeUInt32LE(samples * 2, 40);
  for (let i = 0; i < samples; i++) buf.writeInt16LE(Math.round(Math.sin((i / rate) * 2 * Math.PI * 220) * 8000), 44 + i * 2);
  return buf;
}

const audioTime = (page: Page) => page.evaluate(() => (document.getElementById('audio') as HTMLAudioElement).currentTime);
const isPaused = (page: Page) => page.evaluate(() => (document.getElementById('audio') as HTMLAudioElement).paused);
const setAudioTime = (page: Page, t: number) =>
  page.evaluate((t) => ((document.getElementById('audio') as HTMLAudioElement).currentTime = t), t);

async function openFiles(page: Page) {
  await page.goto('/');
  await page.locator('#docInput').setInputFiles(DOCX);
  await page.locator('#audioInput').setInputFiles({ name: 'recording.wav', mimeType: 'audio/wav', buffer: makeWav() });
  await expect(page.locator('.seg')).toHaveCount(3);
  await expect(page.locator('#durTime')).toHaveText('0:20');
}

/** Read the transcript paragraphs from an exported .docx. */
async function docxParagraphs(path: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(readFileSync(path));
  const xml = await zip.file('word/document2.xml')!.async('string');
  return [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) =>
    [...m[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
  );
}

test('shows the transcript with read-only timestamps and editable speakers and text', async ({ page }) => {
  await openFiles(page);
  await expect(page.locator('#docInfo')).toContainText('Transcript.docx · 3 segments · 2 speakers');
  await expect(page.locator('.stamp')).toHaveText(['00:00:03', '00:00:06', '00:00:08']);
  // Timestamps are buttons (click to seek), not editable fields.
  await expect(page.locator('.stamp').first()).not.toHaveAttribute('contenteditable');
  await expect(page.locator('.seg-text').first()).toHaveValue('Cleanliness is the key to a healthy, happy life.');
  await expect(page.locator('select.speaker').nth(2)).toHaveValue('Speaker 2');
});

test('F1 / F2 skip 5 seconds and F3 toggles play/pause, jumping back 0.5s on pause', async ({ page }) => {
  await openFiles(page);
  await page.locator('.seg-text').first().click(); // shortcuts must work while typing

  await setAudioTime(page, 10);
  await page.keyboard.press('F1');
  expect(await audioTime(page)).toBeCloseTo(5, 1);
  await page.keyboard.press('F2');
  await page.keyboard.press('F2');
  expect(await audioTime(page)).toBeCloseTo(15, 1);

  await setAudioTime(page, 4);
  await page.keyboard.press('F3');
  await expect.poll(() => audioTime(page)).toBeGreaterThan(4.6);
  expect(await isPaused(page)).toBe(false);

  // Record the position at the moment pause() is called, then compare.
  await page.evaluate(() => {
    const audio = document.getElementById('audio') as HTMLAudioElement;
    const pause = audio.pause.bind(audio);
    audio.pause = () => {
      (window as unknown as { pausedAt: number }).pausedAt = audio.currentTime;
      pause();
    };
  });
  await page.keyboard.press('F3');
  expect(await isPaused(page)).toBe(true);
  const pausedAt = await page.evaluate(() => (window as unknown as { pausedAt: number }).pausedAt);
  expect(pausedAt - (await audioTime(page))).toBeCloseTo(0.5, 2);
  // The text box still has focus and nothing was typed into it.
  await expect(page.locator('.seg-text').first()).toBeFocused();
  await expect(page.locator('.seg-text').first()).toHaveValue('Cleanliness is the key to a healthy, happy life.');
});

test('clicking a timestamp plays from there and the playing segment is highlighted', async ({ page }) => {
  await openFiles(page);
  await page.locator('.stamp').nth(1).click();
  await expect.poll(() => audioTime(page)).toBeGreaterThan(6.83);
  expect(await isPaused(page)).toBe(false);
  await expect(page.locator('.seg').nth(1)).toHaveClass(/active/);
  await page.keyboard.press('F3');
});

test('speed can be changed with shortcuts', async ({ page }) => {
  await openFiles(page);
  await expect(page.locator('#rate')).toHaveValue('1');
  await page.keyboard.press('Control+Period');
  await expect(page.locator('#rate')).toHaveValue('1.1');
  await page.keyboard.press('Control+Comma');
  await page.keyboard.press('Control+Comma');
  await expect(page.locator('#rate')).toHaveValue('0.9');
});

test('edits are exported to .docx and autosaved across reloads', async ({ page }) => {
  await openFiles(page);

  const second = page.locator('.seg-text').nth(1);
  await second.fill('Cleanliness, really.');
  await page.locator('select.speaker').nth(0).selectOption('Speaker 2');
  // Alt+1 sets the focused segment to the first speaker.
  await page.locator('.seg-text').nth(2).click();
  await page.keyboard.press('Alt+Digit1');
  await expect(page.locator('select.speaker').nth(2)).toHaveValue('Speaker 1');
  await expect(page.locator('#saveStatus')).toHaveText('Saved in this browser');

  const [download] = await Promise.all([page.waitForEvent('download'), page.locator('#exportBtn').click()]);
  expect(download.suggestedFilename()).toBe('Transcript (edited).docx');
  const paragraphs = await docxParagraphs(await download.path());
  expect(paragraphs).toEqual([
    'Transcript',
    '00:00:03 Speaker 2',
    'Cleanliness is the key to a healthy, happy life.',
    '00:00:06 Speaker 1',
    'Cleanliness, really.',
    '00:00:08 Speaker 1',
    'This is a test.',
  ]);

  // Reload: transcript, edits and audio come back.
  await setAudioTime(page, 7);
  await page.keyboard.press('F3');
  await page.keyboard.press('F3');
  await page.reload();
  await expect(page.locator('.seg')).toHaveCount(3);
  await expect(page.locator('.seg-text').nth(1)).toHaveValue('Cleanliness, really.');
  await expect(page.locator('select.speaker').nth(0)).toHaveValue('Speaker 2');
  await expect(page.locator('#audioName')).toHaveText('recording.wav');
  await expect.poll(() => audioTime(page)).toBeGreaterThan(6);

  // Close clears everything.
  page.once('dialog', (d) => d.accept());
  await page.locator('#closeBtn').click();
  await expect(page.locator('#emptyState')).toBeVisible();
  await page.reload();
  await expect(page.locator('#emptyState')).toBeVisible();
});

test('a new speaker can be added from the dropdown', async ({ page }) => {
  await openFiles(page);
  page.once('dialog', (d) => d.accept('Speaker 3'));
  await page.locator('select.speaker').nth(1).selectOption({ label: 'New speaker…' });
  await expect(page.locator('select.speaker').nth(1)).toHaveValue('Speaker 3');
  await expect(page.locator('select.speaker').nth(0).locator('option')).toHaveText([
    'Speaker 1',
    'Speaker 2',
    'Speaker 3',
    'New speaker…',
  ]);
});

test('shortcuts can be remapped and are remembered', async ({ page }) => {
  await openFiles(page);
  await page.locator('#settingsBtn').click();
  await page.locator('[data-action="back"]').click();
  await page.keyboard.press('KeyJ'); // rejected: would block typing
  await expect(page.locator('#bindingMsg')).toContainText('could no longer type');
  await page.keyboard.press('Control+KeyJ');
  await expect(page.locator('[data-action="back"]')).toHaveText(/J/);
  await page.locator('#settingsDialog button[type="submit"]').click();

  await setAudioTime(page, 12);
  await page.keyboard.press('Control+KeyJ');
  expect(await audioTime(page)).toBeCloseTo(7, 1);
  await page.keyboard.press('F1'); // no longer bound
  expect(await audioTime(page)).toBeCloseTo(7, 1);

  await page.reload();
  await expect(page.locator('#shortcutList')).toContainText(/J/);
});
