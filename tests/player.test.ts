import { describe, expect, it } from 'vitest';
import { Player, clampRate } from '../src/player';

/** Minimal stand-in for an <audio> element. */
class FakeMedia {
  currentTime = 0;
  duration = 60;
  paused = true;
  ended = false;
  playbackRate = 1;
  defaultPlaybackRate = 1;
  src = '';
  play() {
    this.paused = false;
    return Promise.resolve();
  }
  pause() {
    this.paused = true;
  }
  load() {}
  removeAttribute() {}
}

function setup(opts = { skipSeconds: 5, pauseRewindSeconds: 0.5 }) {
  const media = new FakeMedia();
  const player = new Player(media as unknown as HTMLMediaElement, () => opts);
  player.load('blob:test');
  return { media, player };
}

describe('Player', () => {
  it('skips back and forward by the configured amount, clamped to the recording', () => {
    const { media, player } = setup();
    media.currentTime = 10;
    player.back();
    expect(media.currentTime).toBe(5);
    player.back();
    player.back();
    expect(media.currentTime).toBe(0);
    player.forward();
    expect(media.currentTime).toBe(5);
    media.currentTime = 58;
    player.forward();
    expect(media.currentTime).toBe(60);
  });

  it('toggles play/pause and jumps back half a second when pausing', async () => {
    const { media, player } = setup();
    media.currentTime = 12;
    player.toggle();
    await Promise.resolve();
    expect(media.paused).toBe(false);
    media.currentTime = 20;
    player.toggle();
    expect(media.paused).toBe(true);
    expect(media.currentTime).toBeCloseTo(19.5);
    // Resuming doesn't move the position.
    player.toggle();
    expect(media.currentTime).toBeCloseTo(19.5);
  });

  it('does not rewind when already paused or before the start', () => {
    const { media, player } = setup();
    media.currentTime = 0.2;
    media.paused = false;
    player.pause();
    expect(media.currentTime).toBe(0);
    player.pause();
    expect(media.currentTime).toBe(0);
  });

  it('does nothing without audio loaded', () => {
    const media = new FakeMedia();
    const player = new Player(media as unknown as HTMLMediaElement, () => ({ skipSeconds: 5, pauseRewindSeconds: 0.5 }));
    media.currentTime = 10;
    player.back();
    player.toggle();
    expect(media.currentTime).toBe(10);
    expect(media.paused).toBe(true);
  });

  it('changes speed within limits', () => {
    const { media, player } = setup();
    expect(player.changeRate(0.1)).toBe(1.1);
    expect(media.playbackRate).toBe(1.1);
    expect(player.setRate(5)).toBe(2);
    expect(player.setRate(0.1)).toBe(0.5);
    expect(clampRate(0.7000000001)).toBe(0.7);
  });
});
