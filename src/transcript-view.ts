/** Renders the editable transcript: read-only timestamps, speaker dropdowns and text boxes. */
import { newSegmentId } from './docx';
import { activeSegmentIndex, formatStamp } from './time';
import type { Segment } from './types';

export interface TranscriptCallbacks {
  /** Text or speaker changed, or a box was added, split or merged. */
  onEdit(): void;
  /** The timestamp of a segment was clicked. */
  onSeek(segment: Segment): void;
  /** Boxes were added or removed. */
  onStructureChange(): void;
  /** Current audio position and duration, or null when no audio is loaded. */
  audio(): { time: number; duration: number } | null;
}

/** Keep new start times at least this far from their neighbours. */
const MIN_GAP = 0.01;

const NEW_SPEAKER = '\u0000new';
const SPEAKER_COLOURS = 8;
const naturalCompare = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' }).compare;
/** Don't auto-scroll for this long after the user last typed. */
const TYPING_GRACE_MS = 4000;

export class TranscriptView {
  private segments: Segment[] = [];
  private rows: HTMLElement[] = [];
  private starts: number[] = [];
  /** Speakers created with "New speaker…" but not used yet. */
  private extraSpeakers = new Set<string>();
  private activeIndex = -1;
  private lastTime = 0;
  private lastTypedAt = 0;
  private readonly manualAutosize =
    typeof CSS === 'undefined' || typeof CSS.supports !== 'function' || !CSS.supports('field-sizing', 'content');

  constructor(
    private readonly container: HTMLElement,
    private readonly scroller: HTMLElement,
    private readonly callbacks: TranscriptCallbacks,
  ) {
    container.addEventListener('input', (e) => this.onInput(e));
    container.addEventListener('change', (e) => this.onChange(e));
    container.addEventListener('click', (e) => this.onClick(e));
    container.addEventListener('keydown', (e) => this.onKeyDown(e));
    if (this.manualAutosize && typeof ResizeObserver !== 'undefined') {
      let lastWidth = 0;
      new ResizeObserver(([entry]) => {
        const width = Math.round(entry.contentRect.width);
        if (width !== lastWidth) {
          lastWidth = width;
          this.autosizeAll();
        }
      }).observe(container);
    }
  }

  setSegments(segments: Segment[]): void {
    this.segments = segments;
    this.starts = segments.map((s) => s.start);
    this.extraSpeakers.clear();
    this.activeIndex = -1;
    this.container.replaceChildren();
    this.rows = segments.map((seg, i) => this.renderRow(seg, i));
    // Layout: gap, row, gap, row, ..., row, gap. Each gap holds an "add box" button.
    if (segments.length > 0) {
      this.container.append(this.renderGap(), ...this.rows.flatMap((row) => [row, this.renderGap()]));
    }
    this.refreshSpeakers();
    this.autosizeAll();
  }

  /** All speakers, naturally sorted ("Speaker 2" before "Speaker 10"). */
  speakers(): string[] {
    const all = new Set(this.segments.map((s) => s.speaker));
    for (const s of this.extraSpeakers) all.add(s);
    return [...all].sort(naturalCompare);
  }

  /** Highlight the segment playing at `time`, scrolling it into view if `follow` is set. */
  setTime(time: number, follow: boolean): void {
    this.lastTime = time;
    const index = activeSegmentIndex(this.starts, time);
    if (index === this.activeIndex) return;
    this.rows[this.activeIndex]?.classList.remove('active');
    this.activeIndex = index;
    const row = this.rows[index];
    if (!row) return;
    row.classList.add('active');
    if (follow && Date.now() - this.lastTypedAt > TYPING_GRACE_MS && !this.isFullyVisible(row)) {
      row.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
  }

  /**
   * Set the speaker of the segment that has keyboard focus to the n-th speaker (0-based).
   * Returns null if no segment is focused, false if there is no such speaker.
   */
  setSpeakerOfFocused(n: number): boolean | null {
    const row = (document.activeElement as HTMLElement | null)?.closest<HTMLElement>('.seg');
    if (!row || !this.container.contains(row)) return null;
    const speaker = this.speakers()[n];
    if (!speaker) return false;
    this.setSpeaker(Number(row.dataset.index), speaker);
    return true;
  }

  // -------------------------------------------------------------------------------------------
  // Adding, splitting and merging boxes
  // -------------------------------------------------------------------------------------------

  private stampFor(seconds: number): string {
    return formatStamp(seconds, this.segments[0]?.stamp ?? '00:00:00');
  }

  /** Audio position if it lies strictly between `lo` and `hi`. */
  private audioTimeBetween(lo: number, hi: number): number | null {
    const audio = this.callbacks.audio();
    return audio && audio.time > lo + MIN_GAP && audio.time < hi - MIN_GAP ? audio.time : null;
  }

  /** Clamp `t` so that lo < t < hi when there is room. */
  private static clampBetween(t: number, lo: number, hi: number): number {
    if (hi - lo <= 2 * MIN_GAP) return (lo + hi) / 2;
    return Math.min(Math.max(t, lo + MIN_GAP), hi - MIN_GAP);
  }

  private split(index: number, caret: number): void {
    const seg = this.segments[index];
    const next = this.segments[index + 1];
    const before = seg.text.slice(0, caret).replace(/\s+$/, '');
    const after = seg.text.slice(caret).replace(/^\s+/, '');

    // Where the new box starts: the audio position if it's inside this box (you've just heard
    // where the break belongs), otherwise an estimate from where the cursor is in the text.
    const hi = next?.start ?? Infinity;
    let start = this.audioTimeBetween(seg.start, hi);
    if (start === null) {
      const end = seg.end ?? next?.start ?? this.callbacks.audio()?.duration ?? seg.start + 1;
      const fraction = seg.text.length > 0 ? caret / seg.text.length : 0.5;
      start = seg.start + Math.max(0, end - seg.start) * fraction;
    }
    start = TranscriptView.clampBetween(start, seg.start, Number.isFinite(hi) ? hi : start + 2 * MIN_GAP);

    const created: Segment = {
      id: newSegmentId(),
      stamp: this.stampFor(start),
      start,
      end: seg.end !== null && seg.end > start ? seg.end : null,
      speaker: seg.speaker,
      text: after,
    };
    seg.text = before;
    seg.end = start;
    const area = this.textArea(index);
    area.value = before;
    if (this.manualAutosize) this.autosize([area]);

    this.insertAt(index + 1, created);
    this.focusText(index + 1, 0);
  }

  /** Add an empty box before the segment at `at` (or at the end when `at` is the segment count). */
  private addEmpty(at: number): void {
    const prev = this.segments[at - 1];
    const next = this.segments[at];
    const lo = prev?.start ?? 0;
    const duration = this.callbacks.audio()?.duration;
    const hi = next?.start ?? (duration && duration > lo ? duration : lo + 2);

    let start = this.audioTimeBetween(prev ? lo : -MIN_GAP, hi);
    if (start === null) {
      start = prev?.end != null && prev.end > lo && prev.end < hi ? prev.end : (lo + hi) / 2;
    }
    start = prev ? TranscriptView.clampBetween(start, lo, hi) : Math.max(0, Math.min(start, hi - MIN_GAP));

    if (prev && prev.end !== null && prev.end > start) prev.end = start;
    this.insertAt(at, {
      id: newSegmentId(),
      stamp: this.stampFor(start),
      start,
      end: null,
      speaker: prev?.speaker ?? next?.speaker ?? 'Speaker 1',
      text: '',
    });
    this.focusText(at, 0);
  }

  /** Merge the segment at `index` into the one above it. */
  private mergeIntoPrevious(index: number): void {
    const seg = this.segments[index];
    const prev = this.segments[index - 1];
    if (!seg || !prev) return;
    const joiner = prev.text !== '' && seg.text !== '' ? ' ' : '';
    const caret = prev.text.length;
    prev.text = prev.text + joiner + seg.text;
    prev.end = seg.end;
    const area = this.textArea(index - 1);
    area.value = prev.text;
    if (this.manualAutosize) this.autosize([area]);
    this.removeAt(index);
    this.focusText(index - 1, caret);
  }

  private insertAt(at: number, seg: Segment): void {
    const row = this.renderRow(seg, at);
    const gap = this.renderGap();
    // The gap before position `at` is the one just before the row currently at `at` (or the last gap).
    const gapBefore = (this.rows[at]?.previousElementSibling ?? this.container.lastElementChild) as Element;
    gapBefore.after(row, gap);
    this.segments.splice(at, 0, seg);
    this.rows.splice(at, 0, row);
    this.structureChanged();
    this.autosize([row.querySelector('textarea')!]);
  }

  private removeAt(index: number): void {
    const row = this.rows[index];
    row.nextElementSibling?.remove(); // the gap after it
    row.remove();
    this.segments.splice(index, 1);
    this.rows.splice(index, 1);
    this.structureChanged();
  }

  private structureChanged(): void {
    this.rows.forEach((row, i) => (row.dataset.index = String(i)));
    this.starts = this.segments.map((s) => s.start);
    this.rows[this.activeIndex]?.classList.remove('active');
    this.activeIndex = -1;
    this.refreshSpeakers();
    this.setTime(this.lastTime, false);
    this.lastTypedAt = Date.now();
    this.callbacks.onEdit();
    this.callbacks.onStructureChange();
  }

  private textArea(index: number): HTMLTextAreaElement {
    return this.rows[index].querySelector('textarea')!;
  }

  private focusText(index: number, caret: number): void {
    const area = this.textArea(index);
    area.focus({ preventScroll: true });
    area.setSelectionRange(caret, caret);
    if (!this.isFullyVisible(this.rows[index])) this.rows[index].scrollIntoView({ block: 'nearest' });
  }

  private onKeyDown(e: KeyboardEvent): void {
    const area = e.target;
    if (!(area instanceof HTMLTextAreaElement) || e.isComposing) return;
    const index = this.rowIndex(area);
    if (index < 0) return;
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Enter' && mod && !e.altKey && !e.shiftKey) {
      e.preventDefault();
      this.split(index, area.selectionStart);
    } else if (
      e.key === 'Backspace' &&
      !mod && !e.altKey && !e.shiftKey &&
      area.selectionStart === 0 &&
      area.selectionEnd === 0 &&
      this.segments.length > 1
    ) {
      if (index > 0) {
        e.preventDefault();
        this.mergeIntoPrevious(index);
      } else if (this.segments[0].text === '') {
        // An empty first box: just remove it.
        e.preventDefault();
        this.removeAt(0);
        this.focusText(0, 0);
      }
    }
  }

  // -------------------------------------------------------------------------------------------

  private renderGap(): HTMLElement {
    const gap = document.createElement('div');
    gap.className = 'gap';
    const add = document.createElement('button');
    add.type = 'button';
    add.className = 'add-box';
    add.tabIndex = -1;
    add.textContent = '+ Add box';
    add.title = 'Add an empty box here (it starts at the current audio position if that falls here)';
    gap.append(add);
    return gap;
  }

  private renderRow(seg: Segment, index: number): HTMLElement {
    const row = document.createElement('article');
    row.className = 'seg';
    row.dataset.index = String(index);

    const head = document.createElement('div');
    head.className = 'seg-head';

    const stamp = document.createElement('button');
    stamp.type = 'button';
    stamp.className = 'stamp';
    stamp.tabIndex = -1;
    stamp.textContent = seg.stamp;
    stamp.title = `Play from ${seg.stamp}`;

    const select = document.createElement('select');
    select.className = 'speaker';
    select.setAttribute('aria-label', `Speaker at ${seg.stamp}`);

    head.append(stamp, select);

    const text = document.createElement('textarea');
    text.className = 'seg-text';
    text.rows = 1;
    text.spellcheck = true;
    text.value = seg.text;
    text.setAttribute('aria-label', `Text at ${seg.stamp}`);

    row.append(head, text);
    return row;
  }

  private refreshSpeakers(): void {
    const speakers = this.speakers();
    const colour = new Map(speakers.map((s, i) => [s, i % SPEAKER_COLOURS]));
    const options = speakers.map((s) => {
      const o = document.createElement('option');
      o.value = s;
      o.textContent = s;
      return o;
    });
    const newOption = document.createElement('option');
    newOption.value = NEW_SPEAKER;
    newOption.textContent = 'New speaker…';

    this.rows.forEach((row, i) => {
      const seg = this.segments[i];
      const select = row.querySelector('select')!;
      select.replaceChildren(...options.map((o) => o.cloneNode(true)), newOption.cloneNode(true));
      select.value = seg.speaker;
      row.dataset.colour = String(colour.get(seg.speaker) ?? 0);
    });
  }

  private setSpeaker(index: number, speaker: string): void {
    const seg = this.segments[index];
    if (!seg) return;
    if (seg.speaker !== speaker) {
      seg.speaker = speaker;
      this.extraSpeakers.delete(speaker);
      this.callbacks.onEdit();
    }
    this.refreshSpeakers();
  }

  private nextSpeakerName(): string {
    let max = 0;
    for (const s of this.speakers()) {
      const m = /^Speaker (\d+)$/.exec(s);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return `Speaker ${max + 1}`;
  }

  private rowIndex(target: EventTarget | null): number {
    const row = (target as HTMLElement | null)?.closest<HTMLElement>('.seg');
    return row ? Number(row.dataset.index) : -1;
  }

  private onInput(e: Event): void {
    const target = e.target as HTMLElement;
    if (!(target instanceof HTMLTextAreaElement)) return;
    const seg = this.segments[this.rowIndex(target)];
    if (!seg) return;
    seg.text = target.value;
    this.lastTypedAt = Date.now();
    if (this.manualAutosize) this.autosize([target]);
    this.callbacks.onEdit();
  }

  private onChange(e: Event): void {
    const select = e.target as HTMLElement;
    if (!(select instanceof HTMLSelectElement)) return;
    const index = this.rowIndex(select);
    const seg = this.segments[index];
    if (!seg) return;
    if (select.value !== NEW_SPEAKER) {
      this.setSpeaker(index, select.value);
      return;
    }
    const name = window.prompt('Name of the new speaker:', this.nextSpeakerName())?.trim();
    if (name) {
      this.extraSpeakers.add(name);
      this.setSpeaker(index, name);
    } else {
      select.value = seg.speaker;
    }
  }

  private onClick(e: Event): void {
    const add = (e.target as HTMLElement).closest('.add-box');
    if (add) {
      const after = add.parentElement?.nextElementSibling as HTMLElement | null;
      this.addEmpty(after?.classList.contains('seg') ? Number(after.dataset.index) : this.segments.length);
      return;
    }
    const stamp = (e.target as HTMLElement).closest('.stamp');
    if (!stamp) return;
    const seg = this.segments[this.rowIndex(stamp)];
    if (seg) this.callbacks.onSeek(seg);
  }

  private isFullyVisible(row: HTMLElement): boolean {
    const r = row.getBoundingClientRect();
    const p = this.scroller.getBoundingClientRect();
    return r.top >= Math.max(p.top, 0) && r.bottom <= Math.min(p.bottom, window.innerHeight);
  }

  private autosizeAll(): void {
    if (this.manualAutosize) this.autosize(this.rows.map((r) => r.querySelector('textarea')!));
  }

  /** Grow textareas to fit their content (batched reads/writes to avoid layout thrashing). */
  private autosize(areas: HTMLTextAreaElement[]): void {
    for (const a of areas) a.style.height = 'auto';
    const heights = areas.map((a) => a.scrollHeight);
    areas.forEach((a, i) => (a.style.height = `${heights[i]}px`));
  }
}
