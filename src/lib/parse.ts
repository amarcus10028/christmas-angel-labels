import * as XLSX from 'xlsx'
import type { AngelRecord, Issue, ParseResult } from './types'

type Role = 'number' | 'first' | 'last' | 'name' | 'gender' | 'age' | 'gift'

/** Header text → column role. Order matters: first match wins. */
const HEADER_RULES: [Role, RegExp][] = [
  ['gift', /\b(gift|gifts|choice|choices|wish|wishes|request|requests)\b/],
  ['first', /\bfirst\b|\bgiven\b/],
  ['last', /\blast\b|\bsurname\b|\bfamily name\b/],
  ['name', /\bname\b|\brecipient\b|\bchild\b/],
  ['number', /\bangel\b|^#$|^(no|num|number)\.?$/],
  ['gender', /\bgender\b|\bsex\b|^m ?\/ ?f$/],
  ['age', /\bage\b/],
]

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
}

function findColumns(sheet: XLSX.WorkSheet): ColumnMap | null {
  if (!sheet['!ref']) return null
  const range = XLSX.utils.decode_range(sheet['!ref'])
  const lastScan = Math.min(range.e.r, range.s.r + 30)
  for (let r = range.s.r; r <= lastScan; r++) {
    const roles: Partial<Record<Role, number>> = {}
    const gifts: number[] = []
    const giftHeaders: string[] = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const text = cellText(sheet[XLSX.utils.encode_cell({ r, c })])
      const role = classifyHeader(text)
      if (!role) continue
      if (role === 'gift') {
        gifts.push(c)
        giftHeaders.push(text)
      } else if (roles[role] === undefined) roles[role] = c
    }
    const hasName = roles.first !== undefined || roles.name !== undefined
    if (roles.number === undefined || !hasName) continue

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
    return {
      headerRow: r,
      number: roles.number,
      first: roles.first,
      last: roles.last,
      name: roles.first === undefined ? roles.name : undefined,
      gender: roles.gender,
      age: roles.age,
      gifts,
      giftHeaders,
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

  for (const rec of records) {
    if (!rec.age)
      issues.push({ level: 'warning', kind: 'missing-age', row: rec.row, message: `Row ${rec.row}: ${who(rec)} has no age. The age line will be left off.` })
    if (!rec.gender)
      issues.push({ level: 'warning', kind: 'missing-gender', row: rec.row, message: `Row ${rec.row}: ${who(rec)} has no gender. That line will be left off.` })
    if (!rec.gifts.length)
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

  return { records, issues, sheetName, giftColumns: map.giftHeaders }
}
