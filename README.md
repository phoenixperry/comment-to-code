# Comment-to-Code

Code transcripts by commenting on them in Google Drive. Every comment lands in a Google Sheet, and the coded excerpts go onto a FigJam board as stickies for reflexive thematic analysis. The design and its reasoning are in [SPEC.md](SPEC.md).

**For researchers:** see the [user guide (PDF)](docs/Comment-to-Code%20User%20Guide.pdf). It covers setup from the [template Sheet](https://docs.google.com/spreadsheets/d/15pygKZGIGulipCiUQ0BFFjVLE6YvBoRMJ5xwd9jFisc/copy) and the plugin package in `release/`. To rebuild the PDF after editing `docs/user-guide.html`, print it to PDF with headless Chrome.

## How you code

Highlight a passage in a Google Doc or Word file in the watched folder and comment:

```
#belonging #precarity
She frames the studio as the only place she's "allowed" to fail.
```

- `#tags` are codes. They start with a letter and can contain `-` and `_`. Several per comment is fine.
- The rest of the comment is your **memo**, and replies are appended to it.
- A comment with no tag is kept as an uncoded note.
- To rename or merge codes, edit the **Codes** tab (`canonical` column), not your old comments. `code_raw` keeps what you typed and `code` is what it means now, so merges never rewrite history.
- Housekeeping columns in **Codings** are hidden on first setup. Unhide them any time; the script won't hide them again. Deleted rows show greyed out with a line through them.

## Setup

### 1. The Sheet + script (about 10 minutes)

1. Create a Google Sheet, then open **Extensions → Apps Script**.
2. Copy each file from `apps-script/` into the editor (same names; `.gs` files become script files). In **Project Settings**, tick "Show appsscript.json" and paste in `appsscript.json`.
   *Or with clasp:* `clasp clone <scriptId> --rootDir apps-script`, then `clasp push`.
3. Reload the Sheet. A **Coding** menu appears.
4. **Coding → Run self-tests** (authorise when asked). You should see 17/17 passed.
5. **Coding → Set watched folder…** and paste the folder URL.
6. **Coding → Sync now**. Then fill in `participant` on the **Sources** tab.
7. **Coding → Start auto-sync** to sync every 10 minutes.

### 2. The FigJam connection

1. In the Apps Script editor go to **Deploy → New deployment → Web app**, with *Execute as: Me* and *Who has access: Anyone*.
2. In the Sheet: **Coding → FigJam connection details**. Copy the URL and token.
   ⚠️ Anyone with both can read the excerpts. Keep them private and rotate the token if they leak.
3. Build the plugin: `cd figjam-plugin && npm install && npm run build`.
4. In **Figma desktop**, open a FigJam board and go to **Plugins → Development → Import plugin from manifest…**, then pick `figjam-plugin/manifest.json`.
5. Run it, paste the URL and token, and click **Sync from Sheet**.

## On the board

One **Sync** button does both directions: it sends the board to the Sheet first, then pulls anything new.

- **First sync:** one section per code. **Later syncs:** new excerpts arrive in **Unplaced (new)**. Stickies you've moved never move again.
- **Themes are sections you name.** Drag stickies from any codes into a section called, say, "Studio as refuge". Renaming a plugin-made code section also turns it into a theme. **Sections inside sections are sub-themes:** `Studio as refuge › Failure as ritual`.
- **Recode on the board.** Edit the `#tag` at the bottom of a sticky and Sync. The change is written back to the document:
  - your own Google Doc comment is edited in place;
  - someone else's comment gets a reply `↻ #old → #new`, which counts as the recode;
  - a Word comment is edited inside the .docx and saved as a new version (Drive keeps the old one).
- **Subcodes:** `#failure/ritual` is a subcode of `#failure`.
- Excerpts removed from the Sheet are greyed out and marked `[removed]`, never deleted.
- Select a sticky and click **Open source passage** to jump to the comment.

### What lands in the Sheet

| Tab | What it shows |
|---|---|
| **Codings** → `theme` | the theme path of each excerpt's sticky (falls back to `manual_theme` in Codes for excerpts not on the board yet) |
| **Themes** | one row per theme and sub-theme: sticky count, distinct codes, `codes` with counts (most frequent first), participants |
| **Codes** → `board_themes` | where each code's stickies sit, e.g. `Studio as refuge ×3 · (unthemed) ×1` |
| **Board history** | a dated snapshot per Sync: theme, code, participant, excerpt, position. Keep it for your reflexive account of how themes developed. |

## Updating the code

`clasp push` updates the Sheet's script straight away. The clean template has its own script: push to it with `cp .clasp-template.json .clasp.json && clasp push && git checkout -- .clasp.json`, or keep a separate clone. After a plugin change, run `npm run build` and re-zip `manifest.json`, `ui.html` and `dist/code.js` into `release/Comment-to-Code-FigJam-plugin.zip`. The web app is pinned to a version, so after changing `WebApp.gs`, also run `clasp deploy -i <deploymentId>` (the ID is in `clasp deployments`) to keep the same URL.

## Tests

```
cd test && npm install
npm test        # parser + .docx extraction (same cases as Coding → Run self-tests)
npm run sample  # parse sample-data/P99_sample_interview.docx end-to-end
```

Upload `sample-data/P99_sample_interview.docx` to your watched folder to try the whole flow with fictional data.

## Known limits

- Polling, not push: new comments appear within about 10 minutes, or straight away with **Sync now**.
- Word-desktop comments are re-read whenever the file changes. Editing a Word comment's text keeps its row, because rows are keyed on author and time. If Word has stripped author/date info, an edit replaces the row.
- The plugin only works on the current FigJam page.
- One `comments.list` call per file per run is fine for a study-sized corpus (a few hundred files).
