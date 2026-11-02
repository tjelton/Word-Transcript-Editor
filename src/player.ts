/** Thin wrapper around an <audio> element implementing the transcription controls. */

export const MIN_RATE = 0.5;
export const MAX_RATE = 2;
export const RATE_STEP = 0.1;

export function clampRate(rate: number): number {
  return Math.min(MAX_RATE, Math.max(MIN_RATE, Math.round(rate * 100) / 100));
}

export class Player {
  private hasSource = false;

  constructor(
    readonly audio: HTMLMediaElement,
    private readonly options: () => { skipSeconds: number; pauseRewindSeconds: number },
  ) {}

  get loaded(): boolean {
    return this.hasSource;
  }

  get playing(): boolean {
    return this.hasSource && !this.audio.paused && !this.audio.ended;
  }

  load(src: string): void {
    this.audio.src = src;
    this.audio.load();
    this.hasSource = true;
  }

  unload(): void {
    this.audio.pause();
    this.audio.removeAttribute('src');
    this.audio.load();
    this.hasSource = false;
  }

  private get duration(): number {
    return Number.isFinite(this.audio.duration) ? this.audio.duration : Infinity;
  }

  seek(seconds: number): void {
    if (!this.hasSource) return;
    this.audio.currentTime = Math.min(this.duration, Math.max(0, seconds));
  }

  skip(deltaSeconds: number): void {
    this.seek(this.audio.currentTime + deltaSeconds);
  }

  back(): void {
    this.skip(-this.options().skipSeconds);
  }

  forward(): void {
    this.skip(this.options().skipSeconds);
  }

  play(): Promise<void> {
    if (!this.hasSource) return Promise.resolve();
    return this.audio.play().catch(() => undefined);
  }

  /** Pause and step back slightly, so you hear the last few words again when resuming. */
  pause(): void {
    if (!this.playing) return;
    this.audio.pause();
    this.skip(-this.options().pauseRewindSeconds);
  }

  toggle(): void {
    if (this.playing) this.pause();
    else void this.play();
  }

  setRate(rate: number): number {
    const r = clampRate(rate);
    this.audio.playbackRate = r;
    this.audio.defaultPlaybackRate = r;
    return r;
  }

  changeRate(delta: number): number {
    return this.setRate(this.audio.playbackRate + delta);
  }
}
