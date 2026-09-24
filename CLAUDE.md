# Comment-to-Code: notes for Claude

Qualitative coding for reflexive thematic analysis. Researchers comment on transcripts in Drive (`#code` tags). Apps Script collects the comments into a Google Sheet. A FigJam plugin puts one sticky per (excerpt × code) on a shared board, where sections become themes, and edited `#tags` are written back to the documents. The user-facing docs are `README.md` and `docs/Comment-to-Code User Guide.pdf`. `SPEC.md` is the original design and is out of date: it predates themes, recodes and team features.

This repo is **public**. Never commit live project IDs (Sheet, folder, board, deployment IDs, tokens) or anything from the transcripts. Transcripts are anonymised before they reach Drive, so the live web-app feed is fine ethically. Don't re-raise that.

## Layout

| Path | What |
|---|---|
| `apps-script/` | Bound Apps Script (V8), pushed with clasp. All `.gs` files share one global scope. |
| `  Parse.gs` | Pure: tag regex, `parseComment`, recode replies (`↻ #old → #new`), `replaceTagInText`, `commentToRows`. Also runs under Node. |
| `  Docx.gs` | Word comments: `extractDocxComments` via an XML adapter (XmlService or xmldom), `rewriteDocxCommentTag` (byte-level string edit), upload as a new version. |
| `  DriveComments.gs` | Drive API v3 comments → normalised comment objects. |
| `  Sync.gs` | `syncAll` (trigger/menu), `fileComments_`, `syncFiles_` (targeted re-read after a recode). |
| `  Sheet.gs` | Tab/column layout, formulas, upsert, one-time layout (`LAYOUT_V1..V3` flags), copy detection. |
| `  Board.gs` | Plugin POST → recodes written to docs, BoardState, Board history, Recodes log, Themes, Codes!board_themes. |
| `  WebApp.gs` | `doGet` (rows JSON) / `doPost` (board), token check. |
| `  Menu.gs` | Coding menu; auto-sync ownership; connection dialog (finds its own web app via the Apps Script API). |
| `  Tests.gs` | Shared test cases: run from the menu, or `npm test` in `test/`. |
| `figjam-plugin/` | TypeScript plugin (`src/code.ts`, `ui.html`, `manifest.json`), built to `dist/code.js`. |
| `release/` | `Comment-to-Code-FigJam-plugin.zip` for researchers (manifest + ui + dist). |
| `docs/` | `user-guide.html` → PDF. |
| `sample-data/` | Fictional .docx with Word comments plus its generator. |

## Key design rules (keep these)

- **Column ownership.** The script writes only Codings A–M. N–P (participant, code, theme) are ARRAYFORMULAs in row 1, rewritten every sync (a formula that pointed at a missing tab stays `#REF!` and IFERROR hides it). Columns Q+ belong to the researchers. Every string is written through `asText_` (a leading `'`), so transcript text starting with `=`/`-`/digits isn't parsed as a formula.
- **row_key = comment key + `#` + code_raw.** Drive comments use the comment ID. Word comments use `fileId:sha1(author|date)`, which survives re-saves (w:id renumbers on every save).
- **Nothing is deleted.** Removed codes and comments are marked `deleted=TRUE`, and stickies turn grey with `[removed]`.
- **The board belongs to the researchers.** Sync never moves a sticky. New stickies go into "Unplaced (new)". Themes = sections the plugin didn't create (or its code sections after being renamed); nested sections = `A › B`.
- **Sync = send, then pull.** It sends the board first, so a tag edited on a sticky isn't overwritten.
- **Recodes are written back to the source.** Own Google comment → edited in place. Someone else's → `↻` reply naming the researcher. Word → comment XML edited and uploaded as a new version. Each one is logged in the Recodes tab. Rows are rekeyed so the sticky keeps its place.
- **Team use.** Board-wide sync lock (page plugin data, 90 s). One linked board per Sheet (`BOARD_ID`; a mismatch returns `fatal`). `AUTO_SYNC_BY` owner. Colour-by is shared on the page. Copy-pasted stickies are released as plain notes.
- **Copy safety.** `resetIfCopied_` clears FOLDER_ID / WEB_URL / WEB_TOKEN / layout flags when `SHEET_ID` ≠ the active Sheet.

## Workflow

```
cd test && npm install && npm test          # 29 parser/docx tests (keep them passing)
cd figjam-plugin && npm install && npm run build
clasp push --force                          # uses .clasp.json → which script (see below)
clasp deploy -i <deploymentId> --description "…"   # after WebApp/Board changes: keeps the same URL
```

- **Re-zip the plugin** after a plugin change: `manifest.json`, `ui.html` and `dist/code.js` inside a `Comment-to-Code-plugin/` folder → `release/Comment-to-Code-FigJam-plugin.zip`. Check the zip matches the build.
- **Rebuild the guide PDF:** `"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --no-pdf-header-footer --print-to-pdf="docs/Comment-to-Code User Guide.pdf" file://$PWD/docs/user-guide.html`. Each Part should fill whole pages; check with `pdftotext` and fix the contents page numbers.
- **Scope changes** in `appsscript.json` mean every user re-authorises, so warn the owner.

## Finding the live project on a new machine

The IDs aren't in this repo. To get them:
- **Sheet ID:** the `.gsheet` shortcut in Google Drive for desktop is JSON with a `doc_id`. Or open the Sheet; the ID is in the URL.
- **Script ID:** in the Sheet, **Extensions → Apps Script → Project settings → Script ID**. Create `.clasp.json` as `{"scriptId":"…","rootDir":"apps-script"}` (gitignored). `.clasp-template.json` is a committed example pointing at one real project script.
- **Web app deployment:** `clasp deployments` (use the versioned `@N` one, not `@HEAD`).
- **Board and folder:** ask Phoenix, or look in the Sheet (Coding → FigJam connection details shows the linked board; the watched folder is in Script Properties `FOLDER_ID`).
- **Tools:** `npm i -g @google/clasp`, `clasp login` (interactive: the user runs it with `!`), and the Apps Script API switched on at script.google.com/home/usersettings.

## Gotchas already hit

- **UAL Workspace:** `ScriptApp.getService().getUrl()` returns a dead `/a/<domain>/macros/…` URL. The dialog looks up deployments through the Apps Script API instead, and asks for the URL if that fails.
- **"bad token" in the plugin** always means the URL and the token came from different Sheets.
- **`clasp create`** overwrites `appsscript.json` with a default: restore ours before pushing.
- **Unicode regexes (`u` flag):** don't escape `-` outside a character class (see `escapeRe_`).
- **Checkboxes on a whole column** make `getLastRow()` count every row. Only add them to real rows.
- **Drive `quotedFileContent.value`** is HTML-escaped: run it through `decodeEntities`.
- **Development plugins** only run in the Figma desktop app (Mac/Windows). Linux users can edit the board in the browser; someone else runs Sync.
- **Web app POSTs** use `Content-Type: text/plain` to avoid a CORS preflight. The manifest allows script.google.com, *.googleusercontent.com and accounts.google.com.

## Open items

- The README and guide link a live project Sheet as "the template". Make a blank template and swap the link.
- Investigate: a board recode on a Word comment that didn't take (a plain reply was added instead), and a resolved Word comment marked deleted after the .docx was re-saved.
- Offered, not built: participant from the file name or subfolder; start auto-sync automatically on Set watched folder; make `manual_theme` place new stickies in a matching section (or remove it); several boards per Sheet; publish the plugin.
