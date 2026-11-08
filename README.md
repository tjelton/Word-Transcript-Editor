# Transcript Editor

> [https://tjelton.github.io/Word-Transcript-Editor/](https://tjelton.github.io/Word-Transcript-Editor/)

Correct transcripts made by **Microsoft Word's Transcribe** feature while listening to the recording, using the keyboard.

- The transcript is on the left. Timestamps are read-only; you can change the **speaker** of any segment and edit the **text** directly.
- The audio player is on the right.
- Your edits are **exported back to a .docx** in the same layout as Word's, so you can open the result in Word.
- Everything runs in your browser. **Your transcript and audio are never uploaded anywhere.** Your work is autosaved in the browser (IndexedDB), so a refresh or an accidentally closed tab doesn't lose it.

## Keyboard shortcuts

| Key | Action |
| --- | --- |
| <kbd>F1</kbd> | Back 5 seconds |
| <kbd>F2</kbd> | Forward 5 seconds |
| <kbd>F3</kbd> | Play / pause. Pausing jumps back 0.5 seconds so you re-hear the last words when you resume |
| <kbd>Ctrl</kbd>+<kbd>,</kbd> / <kbd>Ctrl</kbd>+<kbd>.</kbd> | Slower / faster playback (0.5×–2×) |
| <kbd>Alt/⌥</kbd>+<kbd>1</kbd>…<kbd>9</kbd> | Set the speaker of the segment you're editing (in the order of the speaker list) |
| <kbd>Ctrl</kbd>+<kbd>Enter</kbd> (<kbd>⌘</kbd>+<kbd>↩</kbd> on a Mac) | Split the box at the cursor |
| <kbd>Backspace</kbd> at the start of a box | Merge it into the box above |
| Click a timestamp | Play from that segment |

Shortcuts work while you're typing in the transcript. You can remap them and change the 5 s / 0.5 s amounts under **Customise…**; your choices are remembered in the browser.

> **Mac users:** by default F1/F2/F3 control screen brightness and Mission Control. Hold <kbd>fn</kbd> while pressing them, or turn on
> *System Settings → Keyboard → Keyboard Shortcuts → Function Keys → "Use F1, F2, etc. keys as standard function keys"*,
> or remap the shortcuts in the app.

### Fixing where boxes start and end

Word sometimes starts a new box in the middle of a sentence, or puts two speakers in one box.

- **Split a box:** put the cursor where the new box should begin and press <kbd>Ctrl</kbd>+<kbd>Enter</kbd> (<kbd>⌘</kbd>+<kbd>↩</kbd> on a Mac). The text after the cursor moves into a new box with the same speaker; press <kbd>Alt/⌥</kbd>+<kbd>1</kbd>…<kbd>9</kbd> to change it. If the audio is paused inside that box, the new box starts at the audio position. Otherwise the start time is estimated from where the cursor is in the text.
- **Add an empty box:** hover between two boxes and click **+ Add box**. It starts at the audio position if that falls between the two boxes, otherwise between them.
- **Merge boxes:** press <kbd>Backspace</kbd> at the very start of a box to join it onto the end of the box above. Backspace in an empty box removes it.

New boxes get a timestamp in the same style as the rest of the document, and they're included in the export.

**Follow audio** (on by default) highlights the segment being played and scrolls it into view. It won't scroll while you're typing.

## Workflow

1. In Word for the web, transcribe the recording and choose **Add to document → With speakers and timestamps**. Then download the document (**File → Save as → Download a copy**).
2. Download the audio recording. Word saves it to the **Transcribed Files** folder in your OneDrive; the .docx only links to it and doesn't contain it.
3. Open the app and drop both files onto the page (or use the buttons).
4. Edit, then click **Export .docx**. This downloads `<name> (edited).docx`.

### What the export changes

The exported document is a copy of the original with the transcript paragraphs rewritten from your edits, keeping the original paragraph and text formatting. Anything before the transcript (such as the "Transcript" heading), styles, and page setup are kept as they were.

Word also stores a hidden copy of the transcription in the document, with exact timings and the link to the OneDrive recording. The export updates that copy to match your edits, and it's what gives the app precise timings for each segment. If a document doesn't have it, the visible timestamps are used instead.

The app reads Word's **Add to document** layout (`00:00:03 Speaker 1`, then the text). It also reads the reverse layout (`Speaker 1  0:03`) and documents where the header and text share a paragraph.

## Deploying

### GitHub Pages (recommended)

The workflow in `.github/workflows/deploy.yml` runs the tests and publishes the app on every push to `main`.

1. Push this repository to GitHub.
2. In the repository, go to **Settings → Pages** and set **Source** to **GitHub Actions** (one-time setup).
3. Push to `main` (or run the workflow from the **Actions** tab). The site appears at `https://<your-username>.github.io/<repo-name>/`.

GitHub Pages for a *private* repository needs a paid plan (GitHub Pro, which is free for students through the GitHub Student Developer Pack). Even then, the published site is public. That's fine here because the site is only the app: your transcripts and audio stay on your computer.

### Running it on your own computer

You need [Node.js](https://nodejs.org/) 20.19 or newer.

```sh
npm install
npm run dev        # development server at http://localhost:5173
# or
npm run build      # produces a static site in dist/
npm run preview    # serves dist/ at http://localhost:4173
```

`dist/` is a plain static site, so you can host it on any static web host.

## Development

```sh
npm test           # unit tests (Vitest): .docx parsing/export, shortcuts, player
npm run test:e2e   # browser tests (Playwright + Chromium); run `npx playwright install chromium` once first
npx tsc --noEmit   # type-check
```

Code layout:

| File | Purpose |
| --- | --- |
| `src/docx.ts` | Read Word transcripts from .docx and write the edited copy |
| `src/transcript-view.ts` | The editable transcript (speaker dropdowns, text boxes, highlighting) |
| `src/player.ts` | Audio controls (skip, pause-with-rewind, speed) |
| `src/shortcuts.ts` | Remappable keyboard shortcuts and settings |
| `src/storage.ts` | Autosave to IndexedDB |
| `src/main.ts` | Wires the page together |

`Transcript.docx` is a sample transcript used by the tests.
