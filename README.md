# Comment-to-Code

Code transcripts by commenting on them in Google Drive. Every comment lands in a Google Sheet, and the coded excerpts go onto a FigJam board as stickies for reflexive thematic analysis. The design and its reasoning are in [SPEC.md](SPEC.md).

**For researchers:** see the [user guide (PDF)](docs/Comment-to-Code%20User%20Guide.pdf). It covers setup from the template Sheet (ask the project lead for the link) and the plugin package in `release/`. To rebuild the PDF after editing `docs/user-guide.html`, print it to PDF with headless Chrome.

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
4. **Coding → Run self-tests** (authorise when asked). You should see 28/28 passed (`npm test` in `test/` runs one extra Node-only check, so it shows 29/29).
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
- One `comments.list` call per file per run is fine for a study-sized corpus (a few hundred files).

## Collaborators

- **Alix Rule**: collaborator. The idea of making tags using a # go into a spreadsheet is hers as are the base features in her tool below (see comparison chart). She is working on the related researh it is being used with.

## Prior work

[frnsys/drive_tagger](https://github.com/frnsys/drive_tagger) by Francis Tseng (May–July 2019) was the first version of Alix Rule's idea: pull `#tags` out of Google Drive comments and collect the tagged text into a spreadsheet. Comment-to-Code (September 2026) is a separate implementation with no shared code. It applies the same idea to reflexive thematic analysis and adds Word files, a FigJam board and write-back.

[GrayAreaorg/drive-tagger-GA](https://github.com/GrayAreaorg/drive-tagger-GA) (September 2026) is a fork of drive_tagger by Barry Threw, Alix Rule and others at Gray Area, published as Text Tagger ([texttagger.com](https://texttagger.com)) for qualitative data analysis. It keeps the original's approach (Google Docs only, every output tab rewritten on each sync) and adds:

- a `sync-all` mode that syncs many projects from a registry Sheet (project, folder, output Sheet, top tags, active, last sync, status, last error), for running unattended on a server;
- tag counts by documents and occurrences, with common tags highlighted and an optional `--top-tags N` for per-tag tabs;
- document → tag and tag → tag co-occurrence edgelists for network analysis, in place of the original's comment-link graphs.

Comment-to-Code shares no code with the fork either. The fork is built for counting and networks across many projects; Comment-to-Code is built for one team's interpretive work: memos, renaming and merging codes, themes on a board, and recodes written back to the documents.

### Timeline

| Date | Version | What changed |
|---|---|---|
| 4 May – 9 Jul 2019 | [drive_tagger](https://github.com/frnsys/drive_tagger) (Francis Tseng, 12 commits) | Tags from Docs comments and replies into a Sheet: tag list, All Tags, one tab per tag, comment-link graphs |
| 25 Aug 2026 | Unpublished draft of drive_tagger (emailed to Phoenix Perry on 27 Aug 2026) | Document → tag and tag → tag graphs replace the comment-link graphs; `--top-tags N` replaces a tab per tag; tabs resize to fit; fuller setup guide |
| 9–10 Sep 2026 | [drive-tagger-GA](https://github.com/GrayAreaorg/drive-tagger-GA) (Barry Threw: "new Alix Rule version") | Occurrence counts, Tag Summary tab, `sync-all` registry mode for many projects |
| 18–19 Sep 2026 | drive-tagger-GA (Alix Rule) | [texttagger.com](https://texttagger.com) site with privacy and terms pages |
| 20 Sep 2026 | drive-tagger-GA | Sorted output, frequent tags highlighted |
| 24 Sep 2026 | Comment-to-Code (Phoenix Perry), first commit | Separate implementation: Apps Script, Docs and Word comments, FigJam board, recodes written back, team sync |

| | drive_tagger (2019) | Comment-to-Code (2026) |
|---|---|---|
| **Runs as** | Local Python CLI (`python main.py sync FOLDER SHEET`) | Apps Script bound to the Sheet: Coding menu, plus auto-sync every 10 minutes |
| **Auth / setup** | Your own Google Cloud OAuth `credentials.json`, with the Drive and Sheets APIs enabled | Authorise the script once inside the Sheet |
| **Sources** | Google Docs only | Google Docs and Word `.docx` comments |
| **Tag syntax** | `#[A-Za-z0-9-_]+`, lowercased | Unicode letters, `-`, `_`, `/` subcodes (`#failure/ritual`), lowercased |
| **Replies** | Tags in replies count as tags | Replies are appended to the memo; `↻ #old → #new` replies recode |
| **Resolved comments** | Skipped | Kept, with a `resolved` flag |
| **Untagged comments** | Ignored | Kept as uncoded notes |
| **Sheet layout** | Tag list with document counts, an "All Tags" tab, one tab per tag, comment→doc and comment→comment reference graphs | Codings (one row per excerpt × code), Sources (participants), Codes (renames and merges via `canonical`), Themes, Board history, Recodes, Log |
| **On each sync** | Clears and rewrites every tab; deletes tabs for tags that have gone | Upserts by a stable row key; removed rows are marked `deleted`, never removed; columns Q+ belong to researchers |
| **Renaming / merging codes** | Edit the comments themselves | Codes tab (`canonical`), so the original `code_raw` is kept |
| **Links back** | Comment URL (`?disco=`), commenter, doc title | Comment link, author, created/modified times, file name, participant |
| **Cross-references** | Graph of comments that link to other docs/comments | Not supported |
| **Visual analysis** | None | FigJam plugin: one sticky per excerpt × code, sections as themes and sub-themes, board ↔ Sheet sync |
| **Write-back** | None (read-only) | Recodes on the board are written back to the source: Google comments edited or replied to, `.docx` saved as a new version, each one logged |
| **Team use** | Single user | Sync lock, auto-sync owner, one linked board per Sheet, copy detection |
| **Tests** | None | Parser and `.docx` tests (Node and in-Sheet) |

### For reflexive thematic analysis

drive_tagger is a general-purpose tagger: its README doesn't mention thematic analysis or any other method. Comment-to-Code was designed for reflexive TA ([Braun & Clarke](https://www.thematicanalysis.net/)), where codes and themes are the researcher's own interpretation and are expected to change as the analysis develops. It has no inter-rater reliability, codebook enforcement or auto-coding, because those belong to coding-reliability approaches.

How each tool fits Braun & Clarke's six phases:

| Phase | drive_tagger | Comment-to-Code |
|---|---|---|
| **1. Familiarisation** | Untagged comments are dropped, so early notes are lost | Comments without tags are kept as uncoded notes; replies are kept as memos |
| **2. Coding** | Tags in comments. Renaming a code means editing every comment. Each sync rewrites the sheet, so earlier codes leave no trace | Tags in comments, with a memo beside each. Codes are renamed and merged in the Codes tab, and `code_raw` keeps what you first wrote. Removed codes are marked, not deleted, so how the coding changed stays visible |
| **3. Generating initial themes** | One tab per tag gathers each code's excerpts; grouping codes into themes happens outside the tool | Excerpt stickies on a FigJam board; you cluster them into named sections by hand, across codes |
| **4. Developing and reviewing themes** | Not supported | Move stickies between themes, nest sub-themes (`A › B`) and recode on the board. Recodes are written back to the transcript, so the data and the themes stay in step. The Themes tab shows each theme's codes and participants, so thin or lopsided themes stand out |
| **5. Refining, defining and naming themes** | Not supported | Rename a section and the new name reaches every excerpt's `theme`. There is no place yet for a written theme definition: keep those in a column you add yourself in Codings or Codes, or in a separate doc |
| **6. Writing up** | The All Tags tab has excerpts with links back to the comments | Every excerpt links back to its comment, with participant and theme, ready for choosing quotes. Each Sync adds a dated snapshot to Board history, and the Recodes tab logs each change of mind. Together they form an audit trail for your reflexive account of how the themes developed |
| **Reflexivity / team** | Records who commented | Keeps every coder's author and replies. A second coder is treated as a sounding board, not a reliability check: their recodes appear as `↻` replies, naming them, beside the original |

## License

[MIT](LICENSE) © 2026 Phoenix Perry. This covers Comment-to-Code only, not the projects listed under Prior work.
