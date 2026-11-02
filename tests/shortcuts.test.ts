import { describe, expect, it } from 'vitest';
import {
  bindingFromEvent,
  defaultSettings,
  findAction,
  formatBinding,
  normaliseSettings,
  validateBinding,
} from '../src/shortcuts';

const press = (code: string, mods: Partial<Record<'ctrlKey' | 'altKey' | 'shiftKey' | 'metaKey', boolean>> = {}) => ({
  code,
  ctrlKey: false,
  altKey: false,
  shiftKey: false,
  metaKey: false,
  ...mods,
});

describe('default shortcuts', () => {
  const settings = defaultSettings();

  it('maps F1/F2/F3 to back/forward/play-pause', () => {
    expect(findAction(settings, press('F1'))).toBe('back');
    expect(findAction(settings, press('F2'))).toBe('forward');
    expect(findAction(settings, press('F3'))).toBe('togglePlay');
  });

  it('requires modifiers to match exactly', () => {
    expect(findAction(settings, press('F1', { shiftKey: true }))).toBeNull();
    expect(findAction(settings, press('Comma', { ctrlKey: true }))).toBe('slower');
    expect(findAction(settings, press('Comma'))).toBeNull();
  });

  it('uses 5 s skips and a 0.5 s rewind on pause', () => {
    expect(settings.skipSeconds).toBe(5);
    expect(settings.pauseRewindSeconds).toBe(0.5);
  });
});

describe('bindingFromEvent', () => {
  it('ignores modifier-only presses', () => {
    expect(bindingFromEvent(press('ShiftLeft', { shiftKey: true }))).toBeNull();
    expect(bindingFromEvent(press('MetaRight', { metaKey: true }))).toBeNull();
  });

  it('records the key and modifiers', () => {
    expect(bindingFromEvent(press('KeyJ', { ctrlKey: true, altKey: true }))).toEqual({
      code: 'KeyJ',
      ctrl: true,
      alt: true,
      shift: false,
      meta: false,
    });
  });
});

describe('validateBinding', () => {
  const b = (code: string, mods: Partial<{ ctrl: boolean; alt: boolean; shift: boolean; meta: boolean }> = {}) => ({
    code,
    ctrl: false,
    alt: false,
    shift: false,
    meta: false,
    ...mods,
  });

  it('accepts F-keys alone and modified letters', () => {
    expect(validateBinding(b('F4'))).toBeNull();
    expect(validateBinding(b('KeyJ', { ctrl: true }))).toBeNull();
    expect(validateBinding(b('KeyK', { alt: true }))).toBeNull();
  });

  it('rejects keys that would stop you typing', () => {
    expect(validateBinding(b('KeyJ'))).toMatch(/could no longer type/);
    expect(validateBinding(b('Space'))).toMatch(/could no longer type/);
    expect(validateBinding(b('KeyJ', { shift: true }))).toMatch(/could no longer type/);
  });

  it('rejects editing shortcuts and the speaker shortcuts', () => {
    expect(validateBinding(b('KeyZ', { meta: true }))).toMatch(/editing/);
    expect(validateBinding(b('KeyC', { ctrl: true }))).toMatch(/editing/);
    expect(validateBinding(b('Digit2', { alt: true }))).toMatch(/speaker/);
  });
});

describe('formatBinding', () => {
  const ctrlComma = { code: 'Comma', ctrl: true, alt: false, shift: false, meta: false };
  it('formats for Mac and Windows', () => {
    expect(formatBinding(ctrlComma, true)).toBe('⌃,');
    expect(formatBinding(ctrlComma, false)).toBe('Ctrl+,');
    expect(formatBinding({ code: 'KeyJ', ctrl: false, alt: true, shift: true, meta: false }, false)).toBe('Alt+Shift+J');
    expect(formatBinding({ code: 'F1', ctrl: false, alt: false, shift: false, meta: false }, true)).toBe('F1');
    expect(formatBinding(null)).toBe('Not set');
  });
});

describe('normaliseSettings', () => {
  it('falls back to defaults for missing or malformed values', () => {
    expect(normaliseSettings(null)).toEqual(defaultSettings());
    const s = normaliseSettings({
      bindings: { back: { code: 'KeyJ', ctrl: true, alt: false, shift: false, meta: false }, forward: 'oops', faster: null },
      skipSeconds: -3,
      pauseRewindSeconds: 1,
      followAudio: false,
    });
    expect(s.bindings.back?.code).toBe('KeyJ');
    expect(s.bindings.forward?.code).toBe('F2');
    expect(s.bindings.faster).toBeNull();
    expect(s.skipSeconds).toBe(5);
    expect(s.pauseRewindSeconds).toBe(1);
    expect(s.followAudio).toBe(false);
  });
});
