/** Keyboard shortcut model: remappable bindings stored in localStorage. */

export type ActionId = 'back' | 'forward' | 'togglePlay' | 'slower' | 'faster';

export interface KeyBinding {
  /** Physical key (KeyboardEvent.code), so Alt/Option combinations work on every layout. */
  code: string;
  ctrl: boolean;
  alt: boolean;
  shift: boolean;
  meta: boolean;
}

export interface Settings {
  bindings: Record<ActionId, KeyBinding | null>;
  skipSeconds: number;
  pauseRewindSeconds: number;
  followAudio: boolean;
}

export const ACTION_IDS: ActionId[] = ['back', 'forward', 'togglePlay', 'slower', 'faster'];

export function actionLabel(id: ActionId, settings: Settings): string {
  switch (id) {
    case 'back':
      return `Back ${settings.skipSeconds}s`;
    case 'forward':
      return `Forward ${settings.skipSeconds}s`;
    case 'togglePlay':
      return settings.pauseRewindSeconds > 0
        ? `Play / pause (pausing jumps back ${settings.pauseRewindSeconds}s)`
        : 'Play / pause';
    case 'slower':
      return 'Slower playback';
    case 'faster':
      return 'Faster playback';
  }
}

const key = (code: string, mods: Partial<Omit<KeyBinding, 'code'>> = {}): KeyBinding => ({
  code,
  ctrl: false,
  alt: false,
  shift: false,
  meta: false,
  ...mods,
});

export function defaultSettings(): Settings {
  return {
    bindings: {
      back: key('F1'),
      forward: key('F2'),
      togglePlay: key('F3'),
      slower: key('Comma', { ctrl: true }),
      faster: key('Period', { ctrl: true }),
    },
    skipSeconds: 5,
    pauseRewindSeconds: 0.5,
    followAudio: true,
  };
}

const MODIFIER_CODES = new Set([
  'ShiftLeft', 'ShiftRight', 'ControlLeft', 'ControlRight', 'AltLeft', 'AltRight',
  'MetaLeft', 'MetaRight', 'OSLeft', 'OSRight', 'CapsLock', 'Fn', 'FnLock',
]);

type KeyLike = Pick<KeyboardEvent, 'code' | 'ctrlKey' | 'altKey' | 'shiftKey' | 'metaKey'>;

/** Build a binding from a key press, or null if only a modifier key was pressed. */
export function bindingFromEvent(e: KeyLike): KeyBinding | null {
  if (!e.code || MODIFIER_CODES.has(e.code)) return null;
  return { code: e.code, ctrl: e.ctrlKey, alt: e.altKey, shift: e.shiftKey, meta: e.metaKey };
}

/** Keys that are safe to use without a modifier because they never type text. */
export function isNonTypingKey(code: string): boolean {
  return /^F([1-9]|1\d|2[0-4])$/.test(code) || ['Pause', 'ScrollLock', 'Insert', 'PrintScreen'].includes(code);
}

const EDITING_LETTERS = ['KeyA', 'KeyC', 'KeyV', 'KeyX', 'KeyZ', 'KeyY', 'KeyQ', 'KeyW', 'KeyR', 'KeyT', 'KeyN'];

/** Why a binding can't be used, or null if it's fine. */
export function validateBinding(b: KeyBinding): string | null {
  if (!(b.ctrl || b.alt || b.meta || isNonTypingKey(b.code))) {
    return 'Add Ctrl, Alt/Option or ⌘ to that key, otherwise you could no longer type it. (F-keys work on their own.)';
  }
  if ((b.ctrl || b.meta) && !b.alt && !b.shift && EDITING_LETTERS.includes(b.code)) {
    return 'That shortcut is used by the browser or for editing text. Pick another one.';
  }
  if (b.alt && !b.ctrl && !b.meta && !b.shift && /^Digit[1-9]$/.test(b.code)) {
    return 'Alt/Option + 1–9 is reserved for changing the speaker.';
  }
  if ((b.ctrl || b.meta) && !b.alt && !b.shift && (b.code === 'Enter' || b.code === 'NumpadEnter')) {
    return 'That shortcut is reserved for splitting a box.';
  }
  return null;
}

export function bindingsEqual(a: KeyBinding | null, b: KeyBinding | null): boolean {
  return !!a && !!b && a.code === b.code && a.ctrl === b.ctrl && a.alt === b.alt && a.shift === b.shift && a.meta === b.meta;
}

export function findAction(settings: Settings, e: KeyLike): ActionId | null {
  const pressed = bindingFromEvent(e);
  return ACTION_IDS.find((id) => bindingsEqual(settings.bindings[id], pressed)) ?? null;
}

export const IS_MAC =
  typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform || navigator.userAgent);

const CODE_LABELS: Record<string, string> = {
  Comma: ',', Period: '.', Slash: '/', Backslash: '\\', Semicolon: ';', Quote: "'",
  BracketLeft: '[', BracketRight: ']', Minus: '-', Equal: '=', Backquote: '`',
  Space: 'Space', Enter: 'Enter', Escape: 'Esc', Backspace: 'Backspace', Tab: 'Tab',
  ArrowLeft: '←', ArrowRight: '→', ArrowUp: '↑', ArrowDown: '↓',
  PageUp: 'Page Up', PageDown: 'Page Down', Home: 'Home', End: 'End', Delete: 'Delete',
};

export function codeLabel(code: string): string {
  if (CODE_LABELS[code]) return CODE_LABELS[code];
  let m = /^Key([A-Z])$/.exec(code);
  if (m) return m[1];
  m = /^Digit(\d)$/.exec(code);
  if (m) return m[1];
  m = /^Numpad(.+)$/.exec(code);
  if (m) return `Num ${m[1]}`;
  return code;
}

export function formatBinding(b: KeyBinding | null, mac = IS_MAC): string {
  if (!b) return 'Not set';
  const k = codeLabel(b.code);
  if (mac) return `${b.ctrl ? '⌃' : ''}${b.alt ? '⌥' : ''}${b.shift ? '⇧' : ''}${b.meta ? '⌘' : ''}${k}`;
  return [b.ctrl && 'Ctrl', b.alt && 'Alt', b.shift && 'Shift', b.meta && 'Win', k].filter(Boolean).join('+');
}

// ---------------------------------------------------------------------------------------------
// Persistence
// ---------------------------------------------------------------------------------------------

const STORAGE_KEY = 'word-transcript-editor.settings.v1';

function isBinding(v: unknown): v is KeyBinding {
  const b = v as Record<string, unknown> | null;
  return !!b && typeof b.code === 'string' && ['ctrl', 'alt', 'shift', 'meta'].every((k) => typeof b[k] === 'boolean');
}

/** Merge stored settings over the defaults, ignoring anything malformed. */
export function normaliseSettings(raw: unknown): Settings {
  const s = defaultSettings();
  const r = (raw ?? {}) as Partial<Record<keyof Settings, unknown>>;
  if (r.bindings && typeof r.bindings === 'object') {
    for (const id of ACTION_IDS) {
      const b = (r.bindings as Record<string, unknown>)[id];
      if (b === null || isBinding(b)) s.bindings[id] = b;
    }
  }
  if (typeof r.skipSeconds === 'number' && r.skipSeconds > 0 && r.skipSeconds <= 600) s.skipSeconds = r.skipSeconds;
  if (typeof r.pauseRewindSeconds === 'number' && r.pauseRewindSeconds >= 0 && r.pauseRewindSeconds <= 60) {
    s.pauseRewindSeconds = r.pauseRewindSeconds;
  }
  if (typeof r.followAudio === 'boolean') s.followAudio = r.followAudio;
  return s;
}

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return normaliseSettings(raw ? JSON.parse(raw) : null);
  } catch {
    return defaultSettings();
  }
}

export function saveSettings(settings: Settings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // Storage unavailable (private mode etc.): settings just won't persist.
  }
}
