/**
 * Autosave: the open transcript (with edits) and the audio file are kept in the browser's
 * IndexedDB so a refresh or accidental close doesn't lose work. Nothing leaves the computer.
 */
import type { HeaderFormat, Segment } from './types';

const DB_NAME = 'word-transcript-editor';
const STORE = 'kv';
const SESSION_KEY = 'session';
const AUDIO_KEY = 'audio';

export interface SavedSession {
  docName: string;
  /** The original .docx, needed to export the edited copy. */
  docBytes: ArrayBuffer;
  segments: Segment[];
  format: HeaderFormat;
  hasStoredTranscription: boolean;
  /** Audio position in seconds. */
  position: number;
  rate: number;
  savedAt: number;
}

export interface SavedAudio {
  name: string;
  blob: Blob;
}

let dbPromise: Promise<IDBDatabase> | null = null;

function openDb(): Promise<IDBDatabase> {
  dbPromise ??= new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('IndexedDB is not available'));
      return;
    }
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  dbPromise.catch(() => {
    dbPromise = null;
  });
  return dbPromise;
}

async function run<T>(mode: IDBTransactionMode, op: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await openDb();
  return new Promise<T>((resolve, reject) => {
    const tx = db.transaction(STORE, mode);
    const req = op(tx.objectStore(STORE));
    tx.oncomplete = () => resolve(req.result);
    tx.onerror = () => reject(tx.error ?? req.error);
    tx.onabort = () => reject(tx.error ?? new Error('Storage transaction aborted'));
  });
}

export const loadSession = () => run<SavedSession | undefined>('readonly', (s) => s.get(SESSION_KEY));
export const saveSession = (session: SavedSession) => run('readwrite', (s) => s.put(session, SESSION_KEY));
export const loadAudio = () => run<SavedAudio | undefined>('readonly', (s) => s.get(AUDIO_KEY));
export const saveAudio = (audio: SavedAudio) => run('readwrite', (s) => s.put(audio, AUDIO_KEY));
export const clearAudio = () => run('readwrite', (s) => s.delete(AUDIO_KEY));

export async function clearAll(): Promise<void> {
  await run('readwrite', (s) => s.clear());
}
