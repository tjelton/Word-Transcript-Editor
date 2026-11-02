/** Renders the editable transcript: read-only timestamps, speaker dropdowns and text boxes. */
import { activeSegmentIndex } from './time';
import type { Segment } from './types';

export interface TranscriptCallbacks {
  /** Text or speaker changed. */
  onEdit(): void;
  /** The timestamp of a segment was clicked. */
  onSeek(segment: Segment): void;
}

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
    this.container.append(...this.rows);
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
