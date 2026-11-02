import { describe, expect, it } from 'vitest';
import { activeSegmentIndex, formatClock, parseStamp } from '../src/time';

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
