import './style.css';
import { DOCX_MIME, editedFileName, exportTranscriptDocx, parseTranscriptDocx } from './docx';
import { MAX_RATE, MIN_RATE, Player, RATE_STEP } from './player';
import {
  ACTION_IDS,
  IS_MAC,
  actionLabel,
  bindingFromEvent,
  bindingsEqual,
  defaultSettings,
  findAction,
  formatBinding,
  loadSettings,
  saveSettings,
  validateBinding,
  type ActionId,
} from './shortcuts';
import * as storage from './storage';
import { formatClock } from './time';
import { TranscriptView } from './transcript-view';
import type { HeaderFormat, Segment } from './types';

// ---------------------------------------------------------------------------------------------
// Elements & state
// ---------------------------------------------------------------------------------------------

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

const els = {
  docInfo: $('docInfo'),
  saveStatus: $('saveStatus'),
  openDocBtn: $<HTMLButtonElement>('openDocBtn'),
  exportBtn: $<HTMLButtonElement>('exportBtn'),
  closeBtn: $<HTMLButtonElement>('closeBtn'),
  docInput: $<HTMLInputElement>('docInput'),
  audioInput: $<HTMLInputElement>('audioInput'),
  transcriptPane: $('transcriptPane'),
  emptyState: $('emptyState'),
  transcript: $('transcript'),
  audio: $<HTMLAudioElement>('audio'),
  audioName: $('audioName'),
  backBtn: $<HTMLButtonElement>('backBtn'),
  playBtn: $<HTMLButtonElement>('playBtn'),
  fwdBtn: $<HTMLButtonElement>('fwdBtn'),
  seek: $<HTMLInputElement>('seek'),
  curTime: $('curTime'),
  durTime: $('durTime'),
  rate: $<HTMLSelectElement>('rate'),
  follow: $<HTMLInputElement>('follow'),
  shortcutList: $('shortcutList'),
  macHint: $('macHint'),
  settingsBtn: $<HTMLButtonElement>('settingsBtn'),
  settingsDialog: $<HTMLDialogElement>('settingsDialog'),
  bindingRows: $('bindingRows'),
  bindingMsg: $('bindingMsg'),
  skipSeconds: $<HTMLInputElement>('skipSeconds'),
  rewindSeconds: $<HTMLInputElement>('rewindSeconds'),
  resetSettings: $<HTMLButtonElement>('resetSettings'),
  dropOverlay: $('dropOverlay'),
  toast: $('toast'),
};

interface OpenDocument {
  name: string;
  bytes: ArrayBuffer;
  segments: Segment[];
  format: HeaderFormat;
  hasStoredTranscription: boolean;
}

let settings = loadSettings();
let current: OpenDocument | null = null;
let audioUrl: string | null = null;
let pendingPosition: number | null = null;
let capturing: ActionId | null = null;

const player = new Player(els.audio, () => settings);

const view = new TranscriptView(els.transcript, els.transcriptPane, {
  onEdit: () => scheduleSave(),
  onSeek: (seg) => {
    if (!requireAudio()) return;
    player.seek(seg.start);
    void player.play();
  },
});

// ---------------------------------------------------------------------------------------------
// Feedback
// ---------------------------------------------------------------------------------------------

let toastTimer: number | undefined;
function toast(message: string, kind: 'info' | 'error' = 'info', ms = kind === 'error' ? 6000 : 2200): void {
  els.toast.textContent = message;
  els.toast.classList.toggle('error', kind === 'error');
  els.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => (els.toast.hidden = true), ms);
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function requireAudio(): boolean {
  if (player.loaded) return true;
  toast('Open the audio recording first: use "Open audio…" or drag the file onto the page.');
  return false;
}

// ---------------------------------------------------------------------------------------------
// Autosave
// ---------------------------------------------------------------------------------------------

let saveTimer: number | undefined;
let lastPositionSave = 0;
/** Edits not yet written to storage. */
let dirty = false;

function setSaveStatus(text: string, error = false): void {
  els.saveStatus.textContent = text;
  els.saveStatus.classList.toggle('error', error);
}

function scheduleSave(): void {
  dirty = true;
  setSaveStatus('Saving…');
  clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => void saveNow(), 500);
}

/** Write the session to storage. `quiet` saves (audio position) only update the status if edits were pending. */
async function saveNow(quiet = false): Promise<void> {
  clearTimeout(saveTimer);
  if (!current) return;
  const hadEdits = dirty;
  dirty = false;
  try {
    await storage.saveSession({
      docName: current.name,
      docBytes: current.bytes,
      segments: current.segments,
      format: current.format,
      hasStoredTranscription: current.hasStoredTranscription,
      position: player.loaded ? els.audio.currentTime : 0,
      rate: els.audio.playbackRate,
      savedAt: Date.now(),
    });
    lastPositionSave = Date.now();
    if (!quiet || hadEdits) setSaveStatus('Saved in this browser');
  } catch (err) {
    console.error(err);
    dirty = dirty || hadEdits;
    setSaveStatus('Autosave failed: export to keep your work', true);
  }
}

// ---------------------------------------------------------------------------------------------
// Opening files
// ---------------------------------------------------------------------------------------------

const AUDIO_EXT = /\.(mp3|m4a|wav|mp4|aac|ogg|oga|opus|webm|flac|wma)$/i;
const isDocx = (f: File) => /\.docx$/i.test(f.name) || f.type === DOCX_MIME;
const isAudio = (f: File) => /^(audio|video)\//.test(f.type) || AUDIO_EXT.test(f.name);

function showDocument(doc: OpenDocument): void {
  current = doc;
  view.setSegments(doc.segments);
  els.emptyState.hidden = true;
  els.transcript.hidden = false;
  els.exportBtn.disabled = false;
  els.closeBtn.hidden = false;
  const speakers = new Set(doc.segments.map((s) => s.speaker)).size;
  els.docInfo.textContent = `${doc.name} · ${doc.segments.length} segments · ${speakers} speaker${speakers === 1 ? '' : 's'}`;
  els.docInfo.title = els.docInfo.textContent;
  view.setTime(els.audio.currentTime, false);
}

async function openDocFile(file: File): Promise<void> {
  if (
    current &&
    !confirm(`Replace "${current.name}" with "${file.name}"?\n\nEdits to the current transcript will be lost unless you've exported them.`)
  ) {
    return;
  }
  try {
    const bytes = await file.arrayBuffer();
    const parsed = await parseTranscriptDocx(bytes);
    showDocument({ name: file.name, bytes, ...parsed });
    els.transcriptPane.scrollTop = 0;
    await saveNow();
    const timingNote = parsed.hasStoredTranscription ? '' : ' (no precise timings found; using the timestamps in the text)';
    toast(`Opened ${file.name}: ${parsed.segments.length} segments${timingNote}`, 'info', 3500);
  } catch (err) {
    toast(`Couldn't open ${file.name}: ${errorMessage(err)}`, 'error');
  }
}

function setAudioName(name: string | null): void {
  els.audioName.textContent = name ?? 'No audio loaded';
  els.audioName.classList.toggle('none', !name);
  els.seek.disabled = !name;
}

// The audio position is also kept in localStorage, which is written synchronously, so it
// survives a reload that happens before the (asynchronous) IndexedDB autosave completes.
const POSITION_KEY = 'word-transcript-editor.position';

function rememberPosition(): void {
  if (!player.loaded || pendingPosition !== null) return;
  try {
    localStorage.setItem(POSITION_KEY, String(els.audio.currentTime));
  } catch {
    // Storage unavailable: the IndexedDB session still holds an approximate position.
  }
}

function recalledPosition(): number | undefined {
  try {
    const value = Number(localStorage.getItem(POSITION_KEY));
    return localStorage.getItem(POSITION_KEY) !== null && Number.isFinite(value) ? value : undefined;
  } catch {
    return undefined;
  }
}

function forgetPosition(): void {
  try {
    localStorage.removeItem(POSITION_KEY);
  } catch {
    // ignore
  }
}

function openAudio(blob: Blob, name: string, opts: { persist: boolean; position?: number }): void {
  if (audioUrl) URL.revokeObjectURL(audioUrl);
  audioUrl = URL.createObjectURL(blob);
  pendingPosition = opts.position ?? null;
  const rate = els.audio.playbackRate;
  player.load(audioUrl);
  player.setRate(rate);
  setAudioName(name);
  if (opts.persist) {
    forgetPosition();
    storage.saveAudio({ name, blob }).catch((err) => {
      console.error(err);
      toast("The audio couldn't be saved in this browser, so you'll need to reopen it after a refresh.", 'error');
    });
  }
}

async function handleFiles(files: Iterable<File>): Promise<void> {
  const list = [...files];
  const doc = list.find(isDocx);
  const audio = list.find((f) => !isDocx(f) && isAudio(f));
  if (doc) await openDocFile(doc);
  if (audio) {
    openAudio(audio, audio.name, { persist: true });
    if (!doc) toast(`Opened ${audio.name}`);
  }
  const ignored = list.filter((f) => f !== doc && f !== audio);
  if (ignored.length) toast(`Ignored ${ignored.map((f) => f.name).join(', ')}: expected a .docx transcript or an audio file.`, 'error');
}

async function restore(): Promise<void> {
  const [session, audio] = await Promise.all([
    storage.loadSession().catch(() => undefined),
    storage.loadAudio().catch(() => undefined),
  ]);
  if (session) {
    showDocument({
      name: session.docName,
      bytes: session.docBytes,
      segments: session.segments,
      format: session.format,
      hasStoredTranscription: session.hasStoredTranscription,
    });
    player.setRate(session.rate || 1);
    setSaveStatus('Saved in this browser');
  }
  if (audio) openAudio(audio.blob, audio.name, { persist: false, position: recalledPosition() ?? session?.position });
  if (session) {
    const when = new Date(session.savedAt).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
    toast(`Restored your work on ${session.docName} (last saved ${when})`, 'info', 3500);
  }
}

// ---------------------------------------------------------------------------------------------
// Export & close
// ---------------------------------------------------------------------------------------------

function download(data: Uint8Array, name: string): void {
  const url = URL.createObjectURL(new Blob([data as BlobPart], { type: DOCX_MIME }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

async function exportDocument(): Promise<void> {
  if (!current) return;
  try {
    const data = await exportTranscriptDocx(current.bytes, current.segments);
    const name = editedFileName(current.name);
    download(data, name);
    toast(`Downloaded ${name}`);
  } catch (err) {
    toast(`Export failed: ${errorMessage(err)}`, 'error');
  }
}

async function closeAll(): Promise<void> {
  if (!confirm('Close this transcript and audio?\n\nYour edits are only stored in this browser, so export first if you need them.')) return;
  clearTimeout(saveTimer);
  current = null;
  dirty = false;
  // Clear storage before updating the UI, so a reload straight after can't restore the old session.
  await storage.clearAll().catch((err) => console.error(err));
  forgetPosition();
  player.unload();
  if (audioUrl) URL.revokeObjectURL(audioUrl);
  audioUrl = null;
  setAudioName(null);
  view.setSegments([]);
  els.transcript.hidden = true;
  els.emptyState.hidden = false;
  els.exportBtn.disabled = true;
  els.closeBtn.hidden = true;
  els.docInfo.textContent = '';
  setSaveStatus('');
}

// ---------------------------------------------------------------------------------------------
// Player UI
// ---------------------------------------------------------------------------------------------

for (let r = MIN_RATE; r <= MAX_RATE + 1e-9; r += RATE_STEP) {
  const rate = Math.round(r * 10) / 10;
  els.rate.add(new Option(`${rate}×`, String(rate)));
}

function runAction(id: ActionId): void {
  if (!requireAudio()) return;
  switch (id) {
    case 'back':
      player.back();
      break;
    case 'forward':
      player.forward();
      break;
    case 'togglePlay':
      player.toggle();
      break;
    case 'slower':
    case 'faster': {
      const rate = player.changeRate(id === 'slower' ? -RATE_STEP : RATE_STEP);
      toast(`Speed ${rate}×`, 'info', 1200);
      break;
    }
  }
}

let seeking = false;

function updateTime(): void {
  const t = els.audio.currentTime;
  els.curTime.textContent = formatClock(t);
  if (!seeking) els.seek.value = String(t);
  view.setTime(t, settings.followAudio);
}

els.audio.addEventListener('loadedmetadata', () => {
  els.seek.max = String(els.audio.duration || 0);
  els.durTime.textContent = formatClock(els.audio.duration);
  if (pendingPosition !== null) {
    player.seek(pendingPosition);
    pendingPosition = null;
  }
  updateTime();
});
els.audio.addEventListener('durationchange', () => {
  els.seek.max = String(els.audio.duration || 0);
  els.durTime.textContent = formatClock(els.audio.duration);
});
els.audio.addEventListener('timeupdate', () => {
  updateTime();
  if (Date.now() - lastPositionSave > 5000) void saveNow(true);
});
els.audio.addEventListener('seeked', () => {
  updateTime();
  rememberPosition();
});
els.audio.addEventListener('play', () => {
  els.playBtn.classList.add('playing');
  els.playBtn.setAttribute('aria-label', 'Pause');
});
els.audio.addEventListener('pause', () => {
  rememberPosition();
  els.playBtn.classList.remove('playing');
  els.playBtn.setAttribute('aria-label', 'Play');
  void saveNow(true);
});
function syncRate(): void {
  els.rate.value = String(Math.round(els.audio.playbackRate * 10) / 10);
}
els.audio.addEventListener('ratechange', syncRate);
els.audio.addEventListener('error', () => {
  if (!player.loaded) return;
  toast("This audio file can't be played by your browser. Try an .mp3, .m4a or .wav version.", 'error');
});

els.playBtn.addEventListener('click', () => requireAudio() && player.toggle());
els.backBtn.addEventListener('click', () => requireAudio() && player.back());
els.fwdBtn.addEventListener('click', () => requireAudio() && player.forward());
els.rate.addEventListener('change', () => {
  player.setRate(Number(els.rate.value));
  void saveNow(true);
});
els.seek.addEventListener('pointerdown', () => (seeking = true));
els.seek.addEventListener('pointerup', () => (seeking = false));
els.seek.addEventListener('input', () => player.seek(Number(els.seek.value)));
els.follow.addEventListener('change', () => {
  settings.followAudio = els.follow.checked;
  saveSettings(settings);
});

// ---------------------------------------------------------------------------------------------
// Shortcuts list & settings dialog
// ---------------------------------------------------------------------------------------------

function kbd(text: string): HTMLElement {
  const k = document.createElement('kbd');
  k.textContent = text;
  return k;
}

function renderShortcuts(): void {
  const items: [Node, string][] = ACTION_IDS.map((id) => [kbd(formatBinding(settings.bindings[id])), actionLabel(id, settings)]);
  items.push([kbd(IS_MAC ? '⌥1 … ⌥9' : 'Alt+1 … 9'), 'Set the speaker of the segment you are editing (in list order)']);
  items.push([document.createTextNode('Click time'), 'Play from that segment']);
  els.shortcutList.replaceChildren(
    ...items.flatMap(([key, label]) => {
      const dt = document.createElement('dt');
      dt.append(key);
      const dd = document.createElement('dd');
      dd.textContent = label;
      return [dt, dd];
    }),
  );
  for (const el of document.querySelectorAll('[data-skip]')) el.textContent = String(settings.skipSeconds);
  els.backBtn.setAttribute('aria-label', `Back ${settings.skipSeconds} seconds`);
  els.fwdBtn.setAttribute('aria-label', `Forward ${settings.skipSeconds} seconds`);
  els.follow.checked = settings.followAudio;
}

function setBindingMessage(text: string, info = false): void {
  els.bindingMsg.textContent = text;
  els.bindingMsg.classList.toggle('info', info);
}

function renderSettings(): void {
  els.bindingRows.replaceChildren(
    ...ACTION_IDS.map((id) => {
      const tr = document.createElement('tr');
      const label = document.createElement('td');
      label.textContent = actionLabel(id, settings);
      const cell = document.createElement('td');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'binding';
      btn.dataset.action = id;
      const binding = settings.bindings[id];
      btn.textContent = capturing === id ? 'Press keys…' : formatBinding(binding);
      btn.classList.toggle('capturing', capturing === id);
      btn.classList.toggle('unset', !binding);
      const clear = document.createElement('button');
      clear.type = 'button';
      clear.className = 'clear-binding ghost';
      clear.dataset.clear = id;
      clear.textContent = '✕';
      clear.title = 'Remove this shortcut';
      clear.disabled = !binding;
      cell.append(btn, clear);
      tr.append(label, cell);
      return tr;
    }),
  );
  els.skipSeconds.value = String(settings.skipSeconds);
  els.rewindSeconds.value = String(settings.pauseRewindSeconds);
}

function settingsChanged(): void {
  saveSettings(settings);
  renderSettings();
  renderShortcuts();
}

els.settingsBtn.addEventListener('click', () => {
  capturing = null;
  setBindingMessage('');
  renderSettings();
  els.settingsDialog.showModal();
});

els.bindingRows.addEventListener('click', (e) => {
  const target = e.target as HTMLElement;
  const action = target.closest<HTMLElement>('[data-action]')?.dataset.action as ActionId | undefined;
  const clear = target.closest<HTMLElement>('[data-clear]')?.dataset.clear as ActionId | undefined;
  if (action) {
    capturing = capturing === action ? null : action;
    setBindingMessage('');
    renderSettings();
    els.bindingRows.querySelector<HTMLElement>(`[data-action="${action}"]`)?.focus();
  } else if (clear) {
    settings.bindings[clear] = null;
    capturing = null;
    settingsChanged();
  }
});

function readNumber(input: HTMLInputElement, min: number, max: number, fallback: number): number {
  const v = Number(input.value);
  return Number.isFinite(v) && v >= min && v <= max ? Math.round(v * 100) / 100 : fallback;
}

els.skipSeconds.addEventListener('change', () => {
  settings.skipSeconds = readNumber(els.skipSeconds, 1, 600, settings.skipSeconds);
  settingsChanged();
});
els.rewindSeconds.addEventListener('change', () => {
  settings.pauseRewindSeconds = readNumber(els.rewindSeconds, 0, 60, settings.pauseRewindSeconds);
  settingsChanged();
});
els.resetSettings.addEventListener('click', () => {
  settings = { ...defaultSettings(), followAudio: settings.followAudio };
  capturing = null;
  setBindingMessage('Shortcuts restored to F1 / F2 / F3.', true);
  settingsChanged();
});
els.settingsDialog.addEventListener('cancel', (e) => {
  if (capturing) e.preventDefault();
});
els.settingsDialog.addEventListener('close', () => {
  capturing = null;
});

function captureBinding(e: KeyboardEvent, action: ActionId): void {
  e.preventDefault();
  e.stopPropagation();
  if (e.code === 'Escape' && !e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey) {
    capturing = null;
    setBindingMessage('');
    renderSettings();
    return;
  }
  const binding = bindingFromEvent(e);
  if (!binding) return; // only a modifier so far
  const problem = validateBinding(binding);
  if (problem) {
    setBindingMessage(problem);
    return;
  }
  let message = '';
  for (const other of ACTION_IDS) {
    if (other !== action && bindingsEqual(settings.bindings[other], binding)) {
      settings.bindings[other] = null;
      message = `${formatBinding(binding)} was removed from “${actionLabel(other, settings)}”.`;
    }
  }
  settings.bindings[action] = binding;
  capturing = null;
  setBindingMessage(message, true);
  settingsChanged();
  els.bindingRows.querySelector<HTMLElement>(`[data-action="${action}"]`)?.focus();
}

// ---------------------------------------------------------------------------------------------
// Global keyboard handling (capture phase, so shortcuts work while typing in the transcript)
// ---------------------------------------------------------------------------------------------

window.addEventListener(
  'keydown',
  (e) => {
    if (capturing) {
      captureBinding(e, capturing);
      return;
    }
    if (e.isComposing) return;

    const action = findAction(settings, e);
    if (action) {
      e.preventDefault();
      e.stopPropagation();
      if (e.repeat && action === 'togglePlay') return;
      runAction(action);
      return;
    }

    if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && /^Digit[1-9]$/.test(e.code)) {
      const n = Number(e.code.slice(5));
      const result = view.setSpeakerOfFocused(n - 1);
      if (result !== null) {
        e.preventDefault();
        if (!result) toast(`There is no speaker number ${n} in the list.`);
      }
    }
  },
  { capture: true },
);

// ---------------------------------------------------------------------------------------------
// File pickers & drag and drop
// ---------------------------------------------------------------------------------------------

els.openDocBtn.addEventListener('click', () => els.docInput.click());
document.querySelectorAll<HTMLElement>('[data-open]').forEach((btn) =>
  btn.addEventListener('click', () => (btn.dataset.open === 'doc' ? els.docInput : els.audioInput).click()),
);
els.docInput.addEventListener('change', () => {
  const file = els.docInput.files?.[0];
  els.docInput.value = '';
  if (file) void openDocFile(file);
});
els.audioInput.addEventListener('change', () => {
  const file = els.audioInput.files?.[0];
  els.audioInput.value = '';
  if (file) void handleFiles([file]);
});
els.exportBtn.addEventListener('click', () => void exportDocument());
els.closeBtn.addEventListener('click', () => void closeAll());

let dragDepth = 0;
const hasFiles = (e: DragEvent) => !!e.dataTransfer && [...e.dataTransfer.types].includes('Files');
window.addEventListener('dragenter', (e) => {
  if (!hasFiles(e)) return;
  dragDepth += 1;
  els.dropOverlay.hidden = false;
});
window.addEventListener('dragleave', (e) => {
  if (!hasFiles(e)) return;
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) els.dropOverlay.hidden = true;
});
window.addEventListener('dragover', (e) => {
  if (hasFiles(e)) e.preventDefault();
});
window.addEventListener('drop', (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  els.dropOverlay.hidden = true;
  void handleFiles(e.dataTransfer!.files);
});

// Flush pending edits when the tab is hidden or closed.
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') {
    rememberPosition();
    void saveNow(true);
  }
});
window.addEventListener('pagehide', rememberPosition);

// ---------------------------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------------------------

els.macHint.hidden = !IS_MAC;
setAudioName(null);
player.setRate(1);
syncRate();
renderShortcuts();
void restore();
