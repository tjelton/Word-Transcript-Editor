import { describe, expect, it } from 'vitest';
import { activeSegmentIndex, formatClock, formatStamp, parseStamp } from '../src/time';

describe('parseStamp', () => {
  it('parses Word timestamp formats', () => {
    expect(parseStamp('00:00:03')).toBe(3);
    expect(parseStamp('01:02:03')).toBe(3723);
    expect(parseStamp('1:05')).toBe(65);
    expect(parseStamp('0:05.5')).toBe(5.5);
    expect(parseStamp('0:05,25')).toBe(5.25);
  });

  it('returns NaN for non-timestamps', () => {
    expect(parseStamp('Speaker 1')).toBeNaN();
    expect(parseStamp('')).toBeNaN();
  });
});

describe('formatClock', () => {
  it('formats short and long durations', () => {
    expect(formatClock(0)).toBe('0:00');
    expect(formatClock(65.9)).toBe('1:05');
    expect(formatClock(3723)).toBe('1:02:03');
    expect(formatClock(NaN)).toBe('0:00');
  });
});

describe('activeSegmentIndex', () => {
  const starts = [3.23, 6.83, 8.11];
  it('finds the segment being played', () => {
    expect(activeSegmentIndex(starts, 0)).toBe(-1);
    expect(activeSegmentIndex(starts, 3.23)).toBe(0);
    expect(activeSegmentIndex(starts, 6.9)).toBe(1);
    expect(activeSegmentIndex(starts, 100)).toBe(2);
    expect(activeSegmentIndex([], 5)).toBe(-1);
  });
});

describe('formatStamp', () => {
  it('matches the style of the document\'s timestamps', () => {
    expect(formatStamp(65.9, '00:00:03')).toBe('00:01:05');
    expect(formatStamp(3725, '00:00:03')).toBe('01:02:05');
    expect(formatStamp(65.9, '0:03')).toBe('1:05');
    expect(formatStamp(65, '00:03')).toBe('01:05');
    expect(formatStamp(3725, '0:03')).toBe('1:02:05');
    expect(formatStamp(-1, '0:03')).toBe('0:00');
  });

  it('round-trips through parseStamp (to the whole second)', () => {
    for (const t of [0, 3.23, 59.99, 61, 3599.5, 7322.4]) {
      expect(parseStamp(formatStamp(t, '00:00:00'))).toBe(Math.floor(t));
    }
  });
});
