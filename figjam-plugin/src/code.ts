// Comment-to-Code: FigJam side.
//
// Board rules (the board belongs to the researcher):
//  - new rows become stickies (first sync: one section per code; later: "Unplaced")
//  - changed rows update text/colour in place; position is never touched
//  - rows gone from the sheet are greyed and marked [removed], never deleted

interface Row {
  row_key: string
  comment_id: string
  file_name: string
  participant: string
  excerpt: string
  code: string
  theme: string
  memo: string
  author: string
  resolved: boolean
  link: string
  colour: string
}

type ColourBy = 'theme' | 'participant' | 'code'

type UiMessage =
  | { type: 'save-settings'; settings: Record<string, string> }
  | { type: 'rows'; rows: Row[]; colourBy: ColourBy }
  | { type: 'collect-board' }
  | { type: 'rekey'; rekeys: Record<string, string> }
  | { type: 'unlock' }

const SETTINGS_KEY = 'comment-to-code-settings'
const UNPLACED = 'Unplaced (new)'
const REMOVED_PREFIX = '[removed] '
const GAP = 24
const PAD = 60
const COLS = 4
const REMOVED_HEX = '#D9D9D9'
const LOCK_MS = 90 * 1000 // a sync that crashed can't block the board for longer than this
const me = () => (figma.currentUser && figma.currentUser.name) || 'someone'

if (figma.command === 'open') {
  openSource()
} else {
  figma.showUI(__html__, { width: 340, height: 460, themeColors: true })
  figma.clientStorage.getAsync(SETTINGS_KEY).then((settings) => {
    // Colour-by lives on the board, shared by everyone, so colours don't flip between people.
    const shared = figma.currentPage.getPluginData('colourBy')
    figma.ui.postMessage({ type: 'settings', settings: Object.assign({}, settings || {}, shared ? { colourBy: shared } : {}) })
  })
  figma.ui.onmessage = async (msg: UiMessage) => {
    try {
      if (msg.type === 'save-settings') {
        await figma.clientStorage.setAsync(SETTINGS_KEY, { url: msg.settings.url, token: msg.settings.token })
        if (msg.settings.colourBy) figma.currentPage.setPluginData('colourBy', msg.settings.colourBy)
      } else if (msg.type === 'rows') {
        const summary = await sync(msg.rows, msg.colourBy)
        unlock()
        figma.ui.postMessage({ type: 'status', text: summary })
        figma.notify(summary)
      } else if (msg.type === 'collect-board') {
        const holder = lockHolder()
        if (holder) {
          figma.ui.postMessage({ type: 'status', text: holder + ' is syncing this board right now. Try again in a minute.' })
          return
        }
        lock()
        figma.ui.postMessage({
          type: 'board',
          stickies: collectBoard(),
          by: me(),
          boardId: (figma.fileKey || figma.root.name) + ':' + figma.currentPage.id,
          boardName: figma.root.name + ' / ' + figma.currentPage.name,
        })
      } else if (msg.type === 'rekey') {
        rekey(msg.rekeys)
      } else if (msg.type === 'unlock') {
        unlock()
      }
    } catch (err) {
      unlock()
      const text = 'Error: ' + (err instanceof Error ? err.message : String(err))
      figma.ui.postMessage({ type: 'status', text })
      figma.notify(text, { error: true })
    }
  }
}

function openSource() {
  const sticky = figma.currentPage.selection.find((n) => n.type === 'STICKY' && n.getPluginData('link'))
  if (sticky) figma.openExternal(sticky.getPluginData('link'))
  else figma.notify('Select a coded sticky first.')
  figma.closePlugin()
}

// ---------- board lock (several people share one board) ----------

function lockHolder(): string | null {
  try {
    const l = JSON.parse(figma.currentPage.getPluginData('syncLock') || 'null')
    return l && l.who !== me() && Date.now() - l.at < LOCK_MS ? l.who : null
  } catch (_) {
    return null
  }
}
function lock() { figma.currentPage.setPluginData('syncLock', JSON.stringify({ who: me(), at: Date.now() })) }
function unlock() { figma.currentPage.setPluginData('syncLock', '') }

// ---------- sync ----------

/**
 * Coded stickies on this page, one per row. A copy-pasted sticky carries the
 * original's plugin data; the copy is released as a plain note the sync ignores.
 */
function pluginStickies(): StickyNode[] {
  const found = figma.currentPage.findAllWithCriteria({ types: ['STICKY'], pluginData: { keys: ['row_key'] } })
  const seen = new Set<string>()
  const out: StickyNode[] = []
  for (const s of found) {
    const key = s.getPluginData('row_key')
    if (!key) continue
    if (seen.has(key)) {
      for (const k of ['row_key', 'link', 'code', 'removed', 'sig']) s.setPluginData(k, '')
      s.setRelaunchData({})
      continue
    }
    seen.add(key)
    out.push(s)
  }
  return out
}

async function sync(rows: Row[], colourBy: ColourBy): Promise<string> {
  const existing = pluginStickies()
  const byKey = new Map(existing.map((s) => [s.getPluginData('row_key'), s] as [string, StickyNode]))
  await loadFonts(existing)

  const live = new Set<string>()
  const fresh: Row[] = []
  let updated = 0
  let removed = 0
  for (const row of rows) {
    live.add(row.row_key)
    const s = byKey.get(row.row_key)
    if (!s) fresh.push(row)
    else if (applyRow(s, row, colourBy)) updated++
  }
  for (const [key, s] of byKey) {
    if (!live.has(key) && s.getPluginData('removed') !== '1') {
      markRemoved(s)
      removed++
    }
  }

  if (fresh.length) {
    const made: StickyNode[] = []
    for (const row of fresh) {
      const s = figma.createSticky()
      await loadFonts([s])
      applyRow(s, row, colourBy)
      made.push(s)
    }
    if (existing.length === 0) layoutByCode(made, fresh)
    else placeUnplaced(made)
    figma.viewport.scrollAndZoomIntoView(made)
  }
  return `${fresh.length} new · ${updated} updated · ${removed} marked removed`
}

/** Set text, colour and metadata. Returns true if anything visible changed. */
function applyRow(s: StickyNode, row: Row, colourBy: ColourBy): boolean {
  const text = stickyText(row)
  const hex = colourFor(row, colourBy)
  const sig = text + '|' + hex
  s.setPluginData('row_key', row.row_key)
  s.setPluginData('link', row.link)
  s.setPluginData('code', row.code)
  s.setPluginData('removed', '0')
  s.setRelaunchData(row.link ? { open: row.file_name } : {})
  s.name = row.code ? '#' + row.code : 'uncoded'
  if (s.getPluginData('sig') === sig) return false
  s.text.characters = text
  s.fills = [{ type: 'SOLID', color: hexToRgb(hex) }]
  s.authorVisible = false
  s.setPluginData('sig', sig)
  return true
}

function markRemoved(s: StickyNode) {
  if (!s.text.characters.startsWith(REMOVED_PREFIX)) s.text.characters = REMOVED_PREFIX + s.text.characters
  s.fills = [{ type: 'SOLID', color: hexToRgb(REMOVED_HEX) }]
  s.setPluginData('removed', '1')
  s.setPluginData('sig', 'removed')
}

function stickyText(row: Row): string {
  const excerpt = clip(row.excerpt, 280)
  const footer = [row.participant || '—', row.code ? '#' + row.code : 'uncoded'].join(' · ')
  const memo = row.memo ? '\n✎ ' + clip(row.memo.split('\n')[0], 140) : ''
  return `“${excerpt}”\n\n${footer}${memo}`
}

function clip(s: string, n: number): string {
  s = (s || '').trim()
  return s.length > n ? s.slice(0, n - 1).replace(/\s+$/, '') + '…' : s
}

// ---------- layout ----------

/** First sync: one section per code, tiled left to right in rows. */
function layoutByCode(stickies: StickyNode[], rows: Row[]) {
  const groups = new Map<string, StickyNode[]>()
  rows.forEach((r, i) => {
    const k = r.code || '(uncoded)'
    if (!groups.has(k)) groups.set(k, [])
    groups.get(k)!.push(stickies[i])
  })
  const names = [...groups.keys()].sort((a, b) => (a === '(uncoded)' ? 1 : b === '(uncoded)' ? -1 : a.localeCompare(b)))
  const origin = figma.viewport.center
  let x = origin.x
  let y = origin.y
  let rowHeight = 0
  let perRow = 0
  const sectionsPerRow = Math.max(2, Math.ceil(Math.sqrt(names.length)))
  for (const name of names) {
    const sec = makeSection(name, groups.get(name)!)
    // Mark as a code section: it only becomes a theme if you rename it.
    sec.setPluginData('role', 'code')
    sec.setPluginData('code', name)
    sec.x = x
    sec.y = y
    rowHeight = Math.max(rowHeight, sec.height)
    x += sec.width + PAD
    if (++perRow === sectionsPerRow) {
      x = origin.x
      y += rowHeight + PAD
      rowHeight = 0
      perRow = 0
    }
  }
}

/** Later syncs: new stickies go to one "Unplaced" section to the right of everything. */
function placeUnplaced(stickies: StickyNode[]) {
  let sec = figma.currentPage.findOne(
    (n) => n.type === 'SECTION' && n.getPluginData('role') === 'unplaced'
  ) as SectionNode | null
  if (!sec) {
    const all = pluginStickies().filter((s) => !stickies.includes(s))
    const right = all.reduce((m, s) => Math.max(m, s.absoluteTransform[0][2] + s.width), figma.viewport.center.x)
    const top = all.reduce((m, s) => Math.min(m, s.absoluteTransform[1][2]), figma.viewport.center.y)
    sec = makeSection(UNPLACED, [])
    sec.setPluginData('role', 'unplaced')
    sec.x = right + PAD * 2
    sec.y = top
  }
  const already = sec.children.filter((n) => n.type === 'STICKY').length
  gridInto(sec, stickies, already)
}

function makeSection(name: string, stickies: StickyNode[]): SectionNode {
  const sec = figma.createSection()
  sec.name = name
  gridInto(sec, stickies, 0)
  return sec
}

/** Append stickies to a section on a COLS-wide grid, starting at slot `start`, and grow it to fit. */
function gridInto(sec: SectionNode, stickies: StickyNode[], start: number) {
  const w = stickies[0] ? stickies[0].width : 240
  const h = stickies[0] ? stickies[0].height : 240
  stickies.forEach((s, i) => {
    const slot = start + i
    sec.appendChild(s)
    s.x = PAD + (slot % COLS) * (w + GAP)
    s.y = PAD * 1.5 + Math.floor(slot / COLS) * (h + GAP)
  })
  const total = Math.max(1, start + stickies.length)
  const cols = Math.min(COLS, total)
  const rows = Math.ceil(total / COLS)
  sec.resizeWithoutConstraints(
    Math.max(sec.width, PAD * 2 + cols * w + (cols - 1) * GAP),
    Math.max(sec.height, PAD * 2.5 + rows * h + (rows - 1) * GAP)
  )
}

// ---------- board → sheet ----------

const THEME_SEP = ' › '
const TAG_RE = /#(\p{L}[\p{L}\p{N}_\/-]*)/u

/** Every live coded sticky: where it sits (theme path), what tag is written on it, and position. */
function collectBoard() {
  const all = pluginStickies()
  const knownCodes = new Set(all.map((s) => s.getPluginData('code')).filter(Boolean))
  return all
    .filter((s) => s.getPluginData('removed') !== '1')
    .map((s) => ({
      row_key: s.getPluginData('row_key'),
      theme: themePath(s, knownCodes),
      tag: tagOnSticky(s),
      x: Math.round(s.absoluteTransform[0][2]),
      y: Math.round(s.absoluteTransform[1][2]),
    }))
}

/**
 * Names of the sections/frames around a sticky, outermost first, skipping the
 * plugin's own ones: "Unplaced" and the per-code sections from the first sync
 * (unless you've renamed them, which turns them into themes).
 */
function themePath(node: SceneNode, knownCodes: Set<string>): string {
  const names: string[] = []
  let p: BaseNode | null = node.parent
  while (p && p.type !== 'PAGE') {
    if (p.type === 'SECTION' || p.type === 'FRAME') {
      const role = p.getPluginData('role')
      const name = p.name.trim()
      const isCodeSection = role === 'code'
        ? name === p.getPluginData('code')
        : !role && knownCodes.has(name.replace(/^#/, '').toLowerCase()) // boards made before roles existed
      if (role !== 'unplaced' && !isCodeSection && name) names.unshift(name)
    }
    p = p.parent
  }
  return names.join(THEME_SEP)
}

/** The code written in the sticky's footer line ("P98 · #privacy"), lower-cased; '' if none. */
function tagOnSticky(s: StickyNode): string {
  const lines = s.text.characters.split('\n').filter((l) => l.trim() && !l.startsWith('✎'))
  const footer = [...lines].reverse().find((l) => l.includes(' · ')) || lines[lines.length - 1] || ''
  const m = TAG_RE.exec(footer.slice(footer.lastIndexOf(' · ') + 1))
  return m ? m[1].replace(/[-_/]+$/, '').toLowerCase() : ''
}

/** After a recode the row has a new key; move the sticky to it so it keeps its place. */
function rekey(rekeys: Record<string, string>) {
  const all = pluginStickies()
  const taken = new Set(all.map((s) => s.getPluginData('row_key')))
  for (const s of all) {
    const to = rekeys[s.getPluginData('row_key')]
    if (!to) continue
    if (taken.has(to)) {
      markRemoved(s) // recoded onto a code this excerpt already has: one sticky is enough
    } else {
      s.setPluginData('row_key', to)
      s.setPluginData('sig', '') // force a redraw on the next pull
      taken.add(to)
    }
  }
}

// ---------- colour & fonts ----------

function colourFor(row: Row, by: ColourBy): string {
  const palette = figma.constants.colors.figJamBaseLight
  const names = Object.keys(palette).filter((k) => !/white|gray|grey/i.test(k))
  if (by === 'code' && row.colour) {
    if (/^#?[0-9a-f]{6}$/i.test(row.colour)) return row.colour.startsWith('#') ? row.colour : '#' + row.colour
    const named = names.find((k) => k.toLowerCase().includes(row.colour.toLowerCase()))
    if (named) return palette[named]
  }
  const key = by === 'theme' ? row.theme || row.code : by === 'participant' ? row.participant : row.code
  if (!key) return palette[Object.keys(palette).find((k) => /gray|grey/i.test(k)) || names[0]]
  return palette[names[hash(key) % names.length]]
}

function hash(s: string): number {
  let h = 2166136261
  for (let i = 0; i < s.length; i++) h = Math.imul(h ^ s.charCodeAt(i), 16777619)
  return h >>> 0
}

function hexToRgb(hex: string): RGB {
  const n = parseInt(hex.replace('#', '').slice(0, 6), 16)
  return { r: ((n >> 16) & 255) / 255, g: ((n >> 8) & 255) / 255, b: (n & 255) / 255 }
}

const loadedFonts = new Set<string>()
async function loadFonts(stickies: StickyNode[]) {
  const fonts: FontName[] = []
  for (const s of stickies) {
    const t = s.text
    if (t.characters.length) fonts.push(...t.getRangeAllFontNames(0, t.characters.length))
    else if (t.fontName !== figma.mixed) fonts.push(t.fontName)
  }
  for (const f of fonts) {
    const id = f.family + '/' + f.style
    if (loadedFonts.has(id)) continue
    await figma.loadFontAsync(f)
    loadedFonts.add(id)
  }
}
