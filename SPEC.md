# Comment-to-Code: a qualitative coding pipeline

Drive comments → Google Sheet → FigJam board, for reflexive thematic analysis.

Status: spec agreed, 24 Sep 2026. Build in progress.

---

## 1. What this is for

You code transcripts the way you already read them: highlight a passage in a Google Doc or Word file and leave a comment. A script collects every comment from a Drive folder into a Sheet, one row per coded excerpt. From the Sheet the codes go onto a FigJam board as stickies, where you cluster them into candidate themes by hand.

The tool should stay out of the way of the analysis. In reflexive TA (Braun & Clarke), codes are meant to shift, merge and get renamed as your understanding develops, and the researcher's interpretation is the point. So the design follows three rules:

- **Commenting is the only coding interface.** No new app to learn, and no coding UI to maintain.
- **Codes are cheap to change later.** Renaming or merging a code happens once in the Sheet, never by editing old comments.
- **The board belongs to you.** A re-sync can add new stickies, but it never moves, deletes or restyles one you've already arranged.

Out of scope on purpose: inter-rater reliability, a codebook enforcer, or auto-coding. Those belong to coding-reliability approaches, not reflexive TA.

> Jamboard was shut down by Google at the end of 2024, so this spec targets **FigJam** (Figma's whiteboard).

---

## 2. System overview

```
 Drive folder (transcripts)
   ├─ Google Docs ──┐
   └─ .docx files ──┤   Apps Script (time trigger, every 10 min)
                    ├─► Drive API: comments.list      ─┐
                    └─► unzip .docx → comments.xml    ─┤
                                                       ▼
                                          Google Sheet
                                          ├─ Codings   (script-owned rows)
                                          ├─ Codes     (you own: rename/merge/theme)
                                          ├─ Sources   (file → participant ID)
                                          └─ Log
                                                       │
                                  CSV export or JSON web app
                                                       ▼
                                 FigJam plugin → stickies in sections
                                                       │  (v3)
                                  sticky positions → themes back to Sheet
```

---

## 3. Part A: Apps Script ingester

### 3.1 Watching the folder

Apps Script has no real "on new comment" event. The options:

| Approach | Verdict |
|---|---|
| **Time-driven trigger polling** every 5–15 min | **Use this.** Simple and reliable. It stays well under the daily trigger quota for a folder of tens to low hundreds of files. |
| Drive push notifications (`changes.watch`) → web app `doPost` | Faster, but channels expire (max ~1 week) and need renewing. You'd still have to fetch the comments afterwards. Not worth it for a coding workflow. |
| Manual "Sync now" menu item | **Also include.** Useful right after a coding session. |

Each run:

1. List the files in the folder (option: recurse into subfolders). Keep Google Docs (`application/vnd.google-apps.document`) and Word files (`.docx`).
2. For each file, fetch comments changed since that file's last sync time, which is stored in `Sources`.
3. Upsert into `Codings`, keyed by comment ID (see 3.4).
4. Record the new sync time. Stop cleanly at ~5 min and resume on the next run, to stay inside the 6-minute execution limit.

### 3.2 Getting comments from Google Docs

Use the **Advanced Drive Service (v3)**: `Drive.Comments.list(fileId, {fields, startModifiedTime, includeDeleted: true})`. Useful fields:

- `id`, `content` (the comment text), `author.displayName`, `createdTime`, `modifiedTime`, `resolved`, `deleted`
- `quotedFileContent.value`: **the highlighted text, which is your excerpt**
- `replies[]`: discussion threads, treated as memos (see 3.3)

Deep link back to the passage: `https://docs.google.com/document/d/{fileId}/edit?disco={commentId}`, which opens the doc on that comment. This matters because it lets you jump from a sticky back to its context.

### 3.3 Getting comments from Word files

This is the tricky part. It depends where the comment was made:

| Where the comment was added | Where it lives | How to read it |
|---|---|---|
| Opened in Drive/Docs (Office editing mode) | Drive comment | Same as 3.2 |
| Made in Word desktop, then uploaded | Inside the file: `word/comments.xml` | Unzip the blob with `Utilities.unzip`, then parse with `XmlService` |

For the second case, the excerpt comes from `word/document.xml`: collect the text runs between `<w:commentRangeStart w:id="N"/>` and `<w:commentRangeEnd w:id="N"/>`. Word comment IDs are small integers that restart with each save, so the upsert key for these is `fileId + hash(author, date, text)`, not `w:id`. Re-uploading a new version of a .docx should update matching rows, not duplicate them.

Recommendation: support both cases, but if you can, code in Google Docs, where it's simpler and you get deep links. The .docx parser is for when collaborators or participants send back Word files.

### 3.4 How a comment becomes a code

Proposed convention, forgiving by design:

```
#belonging #precarity
She frames the studio as the only place she's "allowed" to fail. Link to P02?
```

- Every `#tag` is a code. Tags can contain letters, digits, `-` and `_`. One comment can carry several codes.
- Whatever remains after removing the tags is the **memo**.
- **A comment with no `#tag` still gets ingested**, with `code` left blank. Early on, in the familiarisation phase, you may just want to annotate. Those rows show up in an "uncoded notes" view.
- **Replies** are appended to the memo column, with author and date. They're your analytic conversation with yourself or a collaborator.
- **Resolved** comments are kept, with `resolved = TRUE`. Resolving is a UI gesture, not an analytic decision.
- **Deleted** comments are marked `deleted = TRUE` and hidden by filter, not removed, so there's an audit trail for your reflexive journal.

One row per **(comment × code)**. A comment tagged `#a #b` becomes two rows sharing a `comment_id`. That long ("tidy") shape is what pivot tables, charts and the FigJam plugin all want.

### 3.5 Sheet layout

**`Codings`**: the script writes columns A–M and never touches anything to the right.

| col | field | notes |
|---|---|---|
| A | `row_key` | `commentId#code` |
| B | `comment_id` | Drive comment ID, or `fileId:hash` for Word-embedded comments |
| C | `file_id` | |
| D | `file_name` | hyperlinked to the deep link |
| E | `excerpt` | quoted text |
| F | `code_raw` | exactly as typed, without the `#` |
| G | `memo` | comment text minus tags, plus replies |
| H | `author` | the coder, useful if you work in a pair |
| I | `created` | |
| J | `modified` | |
| K | `resolved` | |
| L | `deleted` | |
| M | `source_type` | `gdoc` / `docx-drive` / `docx-embedded` |
| N | `participant` | **formula**: from `Sources`, so fixing an ID there fixes every row |
| O | `code` | **formula**: canonical code via `Codes` |
| P | `theme` | **formula**: via `Codes` |
| Q+ | yours | free columns for your own notes, never overwritten |

**`Codes`**: generated, then yours. The script appends any `code_raw` it hasn't seen before. Everything else you edit:

| `code_raw` | `canonical` | `theme` | `description` | `colour` | `count` |
|---|---|---|---|---|---|
| belonging | belonging | Studio as refuge | … | yellow | `=COUNTIF(...)` |
| belong | belonging | | | | |
| fail-safe | permission-to-fail | Studio as refuge | | | |

This is how code evolution works: merge `belong` into `belonging` by typing in one cell, and every row follows. Your source comments stay exactly as you wrote them, which is itself a record of how your thinking moved.

**`Sources`**: one row per file, with `file_id`, `file_name`, `participant` (you fill this in, e.g. `P03`), `include` (untick to skip a file), `last_synced`, `comment_rows`, `docx_parsed_at`.

**`Log`**: one line per run (time, files scanned, rows added/updated, errors).

Optional views: a pivot of code × participant (spread), and a filter view of uncoded notes.

### 3.6 Setup UX

A bound script (Extensions → Apps Script) with a custom menu:

- **Coding → Set watched folder…**: paste a folder URL; the script stores the ID in Script Properties.
- **Coding → Sync now**
- **Coding → Start / stop auto-sync**: installs or removes the time trigger.
- **Coding → Export for FigJam**: see 4.2.

The script needs Drive (read) scope plus the sheet. It does not need write access to the documents.

---

## 4. Part B: getting codes into FigJam

The Figma REST API **can't create board content**, so writing stickies means a **FigJam plugin**. There are three levels of effort:

### 4.1 v0: no code

Copy the `excerpt` + `code` columns from a filtered view and paste them into FigJam. It's quick for a first look, but you have to lay things out by hand, re-pasting duplicates stickies, and nothing links back. *(How well FigJam handles pasted spreadsheet cells should be tested before relying on it.)*

### 4.2 v1: small custom FigJam plugin (recommended)

A private plugin (a few hundred lines of TypeScript) that:

- **Reads data one of two ways:**
  1. **Import CSV** (default): pick the file exported by "Export for FigJam". No network, so no participant data leaves Google/Figma through a new channel. This is the safe default for ethics-approved data.
  2. **Live fetch** (opt-in): calls an Apps Script web app (`doGet`) that returns JSON, protected by a secret token. Convenient, but it's a URL that serves transcript excerpts, so check it against your data management plan.
- **Creates one sticky per coding row:**
  - Body: the excerpt (truncated at ~280 characters, with a "…" and the full text in the plugin data)
  - Footer line: `P03 · #belonging`
  - Colour: by **theme**, **participant** or **code**, which you toggle in the plugin UI. FigJam has a small fixed sticky palette (about 10 colours), so colouring by code only works for a small number of codes.
  - Stores `row_key`, `comment_id` and the deep link in `setPluginData` / `setRelaunchData`, so a selected sticky can open its source passage.
- **Layout on first sync:** one FigJam **Section** per canonical code, with stickies in a grid inside it and sections tiled across the canvas. An "Unplaced" section catches new rows on later syncs.
- **Re-sync rules (critical):**
  - A new row gets a new sticky in *Unplaced*.
  - When an existing row changes (edited excerpt, memo or code), update the text and colour **in place**. Never move it.
  - A deleted row gets a strike-through / greyed-out sticky. Never auto-delete.
  - A sticky you've dragged somewhere else stays there.

### 4.3 v2: round trip (themes from the board back to the Sheet)

Once you're clustering, the board holds the analysis. The plugin adds **"Send board to Sheet"**:

- For every plugin-created sticky, it reports which section (or user-drawn frame/group) it's in and its x/y.
- The plugin POSTs to the web app (`doPost`), which writes this to a `Board` sheet: `row_key`, `board_cluster`, `x`, `y`, `snapshot_time`.
- Each send is a dated snapshot, so you can compare how themes looked across sessions. That's useful for the reflexive account of how themes developed (RTA phases 3–5).

---

## 5. Build plan

| Phase | Deliverable | Rough size |
|---|---|---|
| 1 | Apps Script: Google Docs comments → `Codings`, `Codes`, `Sources`, menu, trigger | 1 session |
| 2 | .docx embedded-comment parser | ½–1 session |
| 3 | CSV export + FigJam plugin (import CSV, sections by code, colour toggle, safe re-sync) | 1–2 sessions |
| 4 | Live fetch + round trip to the `Board` sheet | 1 session |

Suggested repo layout:

```
coding_tool/
  SPEC.md
  apps-script/       # clasp project: Code.gs, Drive.gs, Docx.gs, Sheet.gs, WebApp.gs
  figjam-plugin/     # manifest.json (editorType: ["figjam"]), code.ts, ui.html
  sample-data/       # anonymised test transcripts + expected Codings output
```

Use `clasp` to keep the Apps Script in git rather than only in the browser editor.

---

## 6. Risks and edge cases

- **Excerpt drift:** if a transcript is edited after coding, `quotedFileContent` still holds the original quote, which is fine. The deep link may re-anchor or become orphaned. Treat transcripts as frozen once coding starts.
- **Overlapping highlights:** each stays its own comment and row. No merging.
- **Comments on images/tables:** `quotedFileContent` may be empty. Keep the row and flag the excerpt as `[non-text anchor]`.
- **Quota and scale:** one `comments.list` call per changed file per run. Fine for a study-sized corpus (≤ ~300 files). Beyond that, switch to the Drive `changes` feed.
- **Emails:** the Drive API often omits commenter emails, so identify coders by display name.
- **Data protection:** excerpts are participant data. The Sheet inherits the folder's sharing only if you set it up that way, so check it explicitly. Live fetch (4.2) and any published CSV are the main leak risks, which is why both are opt-in.

---

## 7. Decisions

Confirmed (24 Sep 2026):

- **Tag syntax:** `#code` hashtags.
- **Word comments:** both Drive-made and Word-desktop (embedded) comments.
- **FigJam data route:** **live fetch** from a token-protected Apps Script web app, so there's no CSV step. Anyone holding the web-app URL + token can read excerpts. Rotate the token from the menu, and record this route in your data management plan.

Still on defaults (bold). Change any of them:

1. **Subfolders:** **recurse** (e.g. one subfolder per participant) vs. top level only.
2. **Participant ID:** **manual in `Sources`** vs. parsed from the filename (e.g. `P03_interview.docx`).
3. **Solo or team:** if more than one person codes, should the board separate coders (colour by author) or blend them? Reflexive TA usually treats a second coder as a sounding board, not a reliability check, so the default is **blend, with author kept in the data**.
4. **Unit of a sticky:** **one per (excerpt × code)**, so an excerpt with two codes appears twice (standard for affinity mapping) vs. one per excerpt with all its codes listed.
