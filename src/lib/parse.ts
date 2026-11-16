import * as XLSX from 'xlsx'
import type { AngelRecord, Issue, ParseResult } from './types'

type Role = 'number' | 'first' | 'last' | 'name' | 'gender' | 'age' | 'gift' | 'ignore'

/**
 * Header text → column role. Order matters: first match wins.
 * Headings are compared lowercased with extra spaces/underscores removed, so
 * "GEN", "Gen." and "gen" are the same.
 */
const HEADER_RULES: [Role, RegExp][] = [
  ['gift', /\b(gifts?|choices?|wish(es)?|wish ?list|requests?|toys?|presents?|wants?|items?)\b/],
  // Columns that are never printed. Checked before names so "Parent name" isn't read as the child's name.
  ['ignore', /\b(tele?phone|phone|tel|cell|mobile|address|street|apt|city|state|zip|e-?mail|parent|guardian|mother|father|mom|dad|notes?|comments?|table|member|volunteer|donor|sponsor|svdp|date|dob|birth|school|grade|signed|picked|status)\b/],
  ['first', /\bfirst\b|\bgiven\b|^f ?name$/],
  ['last', /\blast\b|\bsurname\b|\bfamily name\b|^l ?name$/],
  ['name', /\bnames?\b|\brecipients?\b|\bchild\b|\bchild's\b|\bkids?\b/],
  ['number', /\bangel\b|^#$|^(no|num|number|id|tag|tag #|ticket|ticket #)\.?$/],
  ['gender', /\bgender\b|\bsex\b|^gen\.?$|^gndr$|^g$|^(m ?\/ ?f|f ?\/ ?m|boy ?\/ ?girl|girl ?\/ ?boy)$/],
  ['age', /\bages?\b|^(yrs?|years?)\.?$|\byears old\b|\(yrs?\)/],
]

const GENDER_VALUE = /^(m|f|male|female|boy|girl|masc|fem)$/i

/** Plain-language name of each field, for showing the volunteer how columns were read. */
const ROLE_LABEL: Record<Exclude<Role, 'ignore'>, string> = {
  number: 'Angel number',
  first: 'First name',
  last: 'Last name',
  name: 'Name',
  gender: 'Gender',
  age: 'Age',
  gift: 'Gift',
}

function norm(s: string): string {
  return s.toLowerCase().replace(/[_\s]+/g, ' ').trim()
}

function classifyHeader(text: string): Role | null {
  const t = norm(text)
  if (!t) return null
  for (const [role, re] of HEADER_RULES) if (re.test(t)) return role
  return null
}

/** Visible text of a cell exactly as Excel shows it, trimmed. Never "null"/"undefined"/"NaN". */
function cellText(cell: XLSX.CellObject | undefined): string {
  if (!cell || cell.v === undefined || cell.v === null) return ''
  let s: string
  if (cell.t === 'n' && typeof cell.v === 'number' && Number.isNaN(cell.v)) return ''
  if (cell.t === 'e') return ''
  if (cell.w !== undefined && cell.w !== null) s = String(cell.w)
  else s = String(cell.v)
  // Keep the text as typed; only line breaks inside a cell become spaces, and the ends are trimmed.
  s = s.replace(/\u00a0/g, ' ').replace(/[ \t]*[\r\n]+[ \t]*/g, ' ').trim()
  if (/^(none|null|undefined|nan|n\/a|#n\/a)$/i.test(s)) return ''
  return s
}

interface ColumnMap {
  headerRow: number // 0-based
  number: number
  first?: number
  last?: number
  name?: number
  gender?: number
  age?: number
  gifts: number[]
  giftHeaders: string[]
  /** How each column was read, left to right. */
  columns: { field: string; header: string; guessed?: boolean }[]
  /** Headings of columns that are not printed (phone, address, notes…). */
  ignored: string[]
}

/** Non-empty values in a column below the header (up to 300 rows). */
function columnValues(sheet: XLSX.WorkSheet, headerRow: number, c: number, lastRow: number): string[] {
  const out: string[] = []
  for (let r = headerRow + 1; r <= Math.min(lastRow, headerRow + 300); r++) {
    const v = cellText(sheet[XLSX.utils.encode_cell({ r, c })])
    if (v) out.push(v)
  }
  return out
}

function mostly(values: string[], test: (v: string) => boolean): boolean {
  return values.length >= 3 && values.filter(test).length / values.length >= 0.9
}

function findColumns(sheet: XLSX.WorkSheet): ColumnMap | null {
  if (!sheet['!ref']) return null
  const range = XLSX.utils.decode_range(sheet['!ref'])
  const lastScan = Math.min(range.e.r, range.s.r + 30)
  for (let r = range.s.r; r <= lastScan; r++) {
    const roles: Partial<Record<Role, number>> = {}
    const gifts: number[] = []
    const giftHeaders: string[] = []
    const headers = new Map<number, string>()
    const unknown: number[] = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const text = cellText(sheet[XLSX.utils.encode_cell({ r, c })])
      headers.set(c, text)
      const role = classifyHeader(text)
      if (!role) {
        unknown.push(c)
        continue
      }
      if (role === 'ignore') continue
      if (role === 'gift') {
        gifts.push(c)
        giftHeaders.push(text)
      } else if (roles[role] === undefined) roles[role] = c
      else unknown.push(c)
    }
    const hasName = roles.first !== undefined || roles.name !== undefined
    if (roles.number === undefined || !hasName) continue

    // No recognizable gender/age heading? Look at the values instead.
    const guessed = new Set<number>()
    if (roles.gender === undefined) {
      const c = unknown.find((c) => mostly(columnValues(sheet, r, c, range.e.r), (v) => GENDER_VALUE.test(v)))
      if (c !== undefined) {
        roles.gender = c
        guessed.add(c)
      }
    }
    if (roles.age === undefined) {
      // Only unlabeled columns: a labeled number column ("Grade", "Table #") is something else.
      const c = unknown.find(
        (c) => !headers.get(c) && c !== roles.gender && mostly(columnValues(sheet, r, c, range.e.r), (v) => /^\d{1,3}$/.test(v) && +v <= 120),
      )
      if (c !== undefined) {
        roles.age = c
        guessed.add(c)
      }
    }

    // Extra unlabeled columns right after the last gift column are also gifts, if they hold anything.
    if (gifts.length) {
      for (let c = gifts[gifts.length - 1] + 1; c <= range.e.c; c++) {
        if (cellText(sheet[XLSX.utils.encode_cell({ r, c })])) break
        let any = false
        for (let rr = r + 1; rr <= range.e.r && !any; rr++) any = !!cellText(sheet[XLSX.utils.encode_cell({ r: rr, c })])
        if (!any) break
        gifts.push(c)
        giftHeaders.push(`Column ${XLSX.utils.encode_col(c)}`)
      }
    }
    const name = roles.first === undefined ? roles.name : undefined
    const used = new Map<number, Exclude<Role, 'ignore'>>()
    const put = (role: Exclude<Role, 'ignore'>, c: number | undefined) => c !== undefined && used.set(c, role)
    put('number', roles.number)
    put('first', roles.first)
    put('last', roles.last)
    put('name', name)
    put('gender', roles.gender)
    put('age', roles.age)
    gifts.forEach((c) => used.set(c, 'gift'))
    const columns = [...used.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([c, role]) => ({ field: ROLE_LABEL[role], header: headers.get(c) || `Column ${XLSX.utils.encode_col(c)}`, guessed: guessed.has(c) || undefined }))
    const ignored = [...headers.entries()]
      .filter(([c, h]) => h && !used.has(c))
      .map(([, h]) => h)

    return {
      headerRow: r,
      number: roles.number,
      first: roles.first,
      last: roles.last,
      name,
      gender: roles.gender,
      age: roles.age,
      gifts,
      giftHeaders,
      columns,
      ignored,
    }
  }
  return null
}

function displayGender(g: string): string {
  const t = g.trim()
  if (/^(m|male|boy|masc)$/i.test(t)) return 'Male'
  if (/^(f|female|girl|fem)$/i.test(t)) return 'Female'
  return t
}

/** Accepts "301", "#301", "Angel #301", "301.0". Returns null if it isn't a whole number. */
function parseAngelNumber(raw: string): { text: string; value: number } | null {
  const m = /^(?:angel\s*)?(?:#|no\.?|number)?\s*(\d+)(?:\.0+)?$/i.exec(raw.trim())
  if (!m) return null
  return { text: m[1], value: parseInt(m[1], 10) }
}

function compressRanges(nums: number[]): string {
  const parts: string[] = []
  for (let i = 0; i < nums.length; i++) {
    let j = i
    while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++
    parts.push(i === j ? `#${nums[i]}` : `#${nums[i]}–#${nums[j]}`)
    i = j
  }
  return parts.join(', ')
}

function who(rec: { number?: string; name?: string }): string {
  if (rec.number && rec.name) return `Angel #${rec.number} (${rec.name})`
  if (rec.number) return `Angel #${rec.number}`
  return rec.name || 'This recipient'
}

export function parseWorkbookData(data: ArrayBuffer | Uint8Array): ParseResult {
  let wb: XLSX.WorkBook
  try {
    wb = XLSX.read(data, { type: 'array', cellDates: false, cellText: true, dense: false })
  } catch {
    return {
      records: [],
      giftColumns: [],
      issues: [{
        level: 'error',
        kind: 'unreadable',
        message: 'This file could not be opened as an Excel spreadsheet. Please choose an .xlsx file saved from Excel.',
      }],
    }
  }
  return parseWorkbook(wb)
}

export function parseWorkbook(wb: XLSX.WorkBook): ParseResult {
  const issues: Issue[] = []
  let map: ColumnMap | null = null
  let sheetName: string | undefined
  for (const name of wb.SheetNames) {
    map = findColumns(wb.Sheets[name])
    if (map) {
      sheetName = name
      break
    }
  }
  if (!map || !sheetName) {
    const empty = wb.SheetNames.every((n) => !wb.Sheets[n]['!ref'])
    issues.push({
      level: 'error',
      kind: empty ? 'no-records' : 'no-header',
      message: empty
        ? 'The spreadsheet is empty — no recipients were found.'
        : 'Could not find the column headings. The spreadsheet needs a heading row with an "Angel #" column and a name column (for example "first name" and "Last name").',
    })
    return { records: [], issues, giftColumns: [] }
  }

  const sheet = wb.Sheets[sheetName]
  const range = XLSX.utils.decode_range(sheet['!ref']!)
  const get = (r: number, c: number | undefined) => (c === undefined ? '' : cellText(sheet[XLSX.utils.encode_cell({ r, c })]))

  const records: AngelRecord[] = []
  const seen = new Map<number, AngelRecord>()

  for (let r = map.headerRow + 1; r <= range.e.r; r++) {
    const row = r + 1
    const numRaw = get(r, map.number)
    const name =
      map.first !== undefined
        ? [get(r, map.first), get(r, map.last)].filter(Boolean).join(' ')
        : get(r, map.name)
    const gender = displayGender(get(r, map.gender))
    const age = get(r, map.age)
    const gifts = map.gifts.map((c) => get(r, c)).filter(Boolean)

    const filled = [numRaw, name, gender, age, ...gifts].filter(Boolean).length
    if (filled === 0) continue // blank row

    const num = numRaw ? parseAngelNumber(numRaw) : null
    // A lone piece of text (a heading, total or note) is not a recipient.
    if (filled === 1 && !num) continue
    if (filled === 1 && num && !name) {
      issues.push({ level: 'error', kind: 'missing-name', row, message: `Row ${row}: Angel #${num.text} is missing a name.` })
      continue
    }

    if (numRaw && !num) {
      issues.push({
        level: 'error',
        kind: 'bad-number',
        row,
        message: `Row ${row}: the angel number “${numRaw}”${name ? ` for ${name}` : ''} is not a number.`,
      })
      continue
    }
    if (!num) {
      issues.push({
        level: 'error',
        kind: 'missing-number',
        row,
        message: `Row ${row}: ${name || 'a recipient'} is missing an angel number.`,
      })
      continue
    }
    if (!name) {
      issues.push({ level: 'error', kind: 'missing-name', row, message: `Row ${row}: Angel #${num.text} is missing a name.` })
      continue
    }

    const rec: AngelRecord = { number: num.text, numberValue: num.value, name, gender, age, gifts, row }
    const prev = seen.get(num.value)
    if (prev) {
      issues.push({
        level: 'error',
        kind: 'duplicate',
        row,
        message: `Row ${row}: Angel #${num.text} is used twice — ${prev.name} (row ${prev.row}) and ${name} (row ${row}). Each angel needs its own number.`,
      })
      continue
    }
    seen.set(num.value, rec)
    records.push(rec)
  }

  records.sort((a, b) => a.numberValue - b.numberValue)

  if (records.length === 0 && !issues.some((i) => i.level === 'error')) {
    issues.push({ level: 'error', kind: 'no-records', message: 'No recipients were found below the heading row.' })
  }

  // A whole missing column gets one note, not one warning per angel.
  if (records.length && map.gender === undefined)
    issues.push({ level: 'info', kind: 'note', message: 'No gender column was found, so the labels won’t show gender. If there is one, give it the heading “gender”.' })
  if (records.length && map.age === undefined)
    issues.push({ level: 'info', kind: 'note', message: 'No age column was found, so the labels won’t show ages. If there is one, give it the heading “age”.' })
  if (records.length && !map.gifts.length)
    issues.push({ level: 'warning', kind: 'note', message: 'No gift columns were found, so every label will say “Gift choices not provided.” Gift column headings need the word “gift”, such as “first gift”.' })

  for (const rec of records) {
    if (!rec.age && map.age !== undefined)
      issues.push({ level: 'warning', kind: 'missing-age', row: rec.row, message: `Row ${rec.row}: ${who(rec)} has no age. The age line will be left off.` })
    if (!rec.gender && map.gender !== undefined)
      issues.push({ level: 'warning', kind: 'missing-gender', row: rec.row, message: `Row ${rec.row}: ${who(rec)} has no gender. That line will be left off.` })
    if (!rec.gifts.length && map.gifts.length)
      issues.push({ level: 'warning', kind: 'no-gifts', row: rec.row, message: `Row ${rec.row}: ${who(rec)} has no gift choices. The label will say “Gift choices not provided.”` })
  }

  // Gaps are informational only: skipped numbers are often intentional.
  const missing: number[] = []
  for (let i = 1; i < records.length; i++) {
    for (let n = records[i - 1].numberValue + 1; n < records[i].numberValue && missing.length <= 500; n++) missing.push(n)
  }
  if (missing.length) {
    const list = missing.length <= 30 ? compressRanges(missing) : `${missing.length} numbers`
    issues.push({
      level: 'info',
      kind: 'gap',
      message: `Not used in this spreadsheet: ${list}. That’s fine if it’s on purpose — no labels are made for unused numbers.`,
    })
  }

  return { records, issues, sheetName, giftColumns: map.giftHeaders, columns: map.columns, ignoredColumns: map.ignored }
}
