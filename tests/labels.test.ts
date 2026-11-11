import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { parseWorkbookData } from '../src/lib/parse'
import { generateDocxBuffer, fitIssues } from '../src/lib/docx'
import { formatDeadline, outputFilename } from '../src/lib/date'
import { getLayout, LAYOUTS, labelsPerPage, pageMargins, BOTTOM_BUFFER } from '../src/lib/layouts'
import { DEFAULT_PARISH, type AngelRecord } from '../src/lib/types'
import { NO_GIFTS_TEXT } from '../src/lib/label'
import { extractDocx, HEADER, referenceRows, workbookBytes, type Row } from './helpers'

const DEC6 = 'Sunday, December 6, 2026'
const DEC13 = 'Sunday, December 13, 2026'
const layout = getLayout('avery-5168')

/** The real parish spreadsheet is never committed (it has children's names). Drop a copy here to test against it. */
const REAL = resolve(__dirname, '../fixtures-private/OLL.xlsx')
const sources: [string, () => Uint8Array][] = [['synthetic reference', () => workbookBytes(referenceRows())]]
if (existsSync(REAL)) sources.push(['real OLL.xlsx', () => new Uint8Array(readFileSync(REAL))])

async function generate(records: AngelRecord[], deadline: string, layoutId = 'avery-5168') {
  const bytes = await generateDocxBuffer({ records, parish: DEFAULT_PARISH, deadlineText: deadline, layout: getLayout(layoutId) })
  return extractDocx(bytes)
}

function expectLabelsMatch(records: AngelRecord[], labels: Awaited<ReturnType<typeof generate>>['labels'], deadline: string) {
  expect(labels.length).toBe(records.length)
  labels.forEach((l, i) => {
    const r = records[i]
    expect(l.number).toBe(r.number)
    expect(l.name).toBe(r.name)
    expect(l.gender).toBe(r.gender)
    expect(l.age).toBe(r.age)
    expect(l.gifts).toEqual(r.gifts.length ? r.gifts : [NO_GIFTS_TEXT])
    expect(l.deadlines).toEqual([`Return deadline: ${deadline}`])
    for (const line of l.lines) expect(line).not.toMatch(/\b(None|null|undefined|NaN)\b/)
  })
}

describe('date formatting', () => {
  it('adds the weekday', () => {
    expect(formatDeadline('2026-12-06')).toBe(DEC6)
    expect(formatDeadline('2026-12-13')).toBe(DEC13)
    expect(formatDeadline('2027-01-01')).toBe('Friday, January 1, 2027')
    expect(formatDeadline('')).toBeNull()
    expect(outputFilename('2026-12-06')).toBe('Christmas_Angel_Labels_2026.docx')
  })
})

describe.each(sources)('%s', (_name, load) => {
  const result = parseWorkbookData(load())

  it('imports all 199 recipients with no blocking errors', () => {
    expect(result.issues.filter((i) => i.level === 'error')).toEqual([])
    expect(result.records.length).toBe(199)
    expect(result.records[0].number).toBe('301')
    expect(result.records.at(-1)!.number).toBe('500')
    expect(result.giftColumns).toEqual(['first gift', 'second gift', 'Third Gift'])
  })

  it('Test A — December 6: every label matches the spreadsheet', async () => {
    const doc = await generate(result.records, DEC6)
    expectLabelsMatch(result.records, doc.labels, DEC6)
    expect(doc.documentXml).not.toContain('December 13')
    expect(doc.emptySlots).toBe(1) // 199 labels → 50 pages, one blank position
  })

  it('Test B — December 13 from the same parsed data', async () => {
    const doc = await generate(result.records, DEC13)
    expectLabelsMatch(result.records, doc.labels, DEC13)
    expect(doc.documentXml).not.toContain('December 6')
  })

  it('Test C — #379–#383 have no age line and a warning each', async () => {
    const doc = await generate(result.records, DEC6)
    for (const n of ['379', '380', '381', '382', '383']) {
      const l = doc.labels.find((x) => x.number === n)!
      expect(l.age).toBe('')
      expect(l.lines.some((x) => x.startsWith('Age'))).toBe(false)
    }
    const ages = result.issues.filter((i) => i.kind === 'missing-age')
    expect(ages.length).toBe(5)
    expect(ages.every((i) => i.level === 'warning')).toBe(true)
  })

  it('Test D — #429 is skipped, not created; #430 follows #428', () => {
    const nums = result.records.map((r) => r.number)
    expect(nums).not.toContain('429')
    expect(nums[nums.indexOf('428') + 1]).toBe('430')
    const gap = result.issues.find((i) => i.kind === 'gap')!
    expect(gap.level).toBe('info')
    expect(gap.message).toContain('#429')
  })

  it('expands M/F (any case) and keeps names trimmed', () => {
    expect(new Set(result.records.map((r) => r.gender))).toEqual(new Set(['Male', 'Female']))
    for (const r of result.records) expect(r.name).toBe(r.name.trim())
  })
})

describe('Test E — error handling', () => {
  const base = () => referenceRows()
  const parse = (rows: Row[]) => parseWorkbookData(workbookBytes(rows))
  const errors = (rows: Row[]) => parse(rows).issues.filter((i) => i.level === 'error')

  it('duplicate angel number', () => {
    const rows = base()
    rows[5][0] = 301
    const e = errors(rows)
    expect(e.length).toBe(1)
    expect(e[0].kind).toBe('duplicate')
    expect(e[0].message).toMatch(/^Row 6: Angel #301 is used twice/)
  })

  it('missing name', () => {
    const rows = base()
    rows[36][1] = null
    rows[36][2] = null
    const e = errors(rows)
    expect(e.map((x) => x.message)).toEqual([`Row 37: Angel #${rows[36][0]} is missing a name.`])
  })

  it('missing angel number', () => {
    const rows = base()
    rows[10][0] = null
    const e = errors(rows)
    expect(e.length).toBe(1)
    expect(e[0].kind).toBe('missing-number')
    expect(e[0].message).toMatch(/^Row 11: .+ is missing an angel number\.$/)
  })

  it('empty worksheet', () => {
    const e = errors([])
    expect(e.length).toBe(1)
    expect(e[0].message).toMatch(/empty|no recipients/i)
  })

  it('headings only, no recipients', () => {
    expect(errors([HEADER])[0].kind).toBe('no-records')
  })

  it('unreadable file', () => {
    const r = parseWorkbookData(new TextEncoder().encode('\u0000\u0001 not a spreadsheet \u0002'))
    expect(r.issues.some((i) => i.level === 'error')).toBe(true)
  })

  it('missing gift choices → warning and placeholder text', async () => {
    const rows = base()
    rows[1].splice(7, 3, null, null, null)
    const r = parse(rows)
    expect(r.issues.filter((i) => i.level === 'error')).toEqual([])
    expect(r.issues.find((i) => i.kind === 'no-gifts')!.message).toMatch(/^Row 2: Angel #301/)
    const doc = await generate(r.records, DEC6)
    expect(doc.labels[0].gifts).toEqual([NO_GIFTS_TEXT])
  })

  it('long name and long gifts shrink to fit instead of overflowing', async () => {
    const rows = base()
    rows[1][1] = 'Maria Guadalupe Esperanza de los Angeles'
    rows[1][2] = 'Rodriguez-Fernandez Villanueva'
    rows[1][7] = 'Brandy Melville zip-up hoodie, medium, navy or green or gray, no logos please, women’s fit'
    rows[1][8] = 'Winter coat with hood, black, women’s size small petite, warm lining, long length below knee'
    rows[1][9] = 'Art supplies: watercolor paints, brushes, sketchbook (spiral, 9×12), colored pencils 24-pack'
    const r = parse(rows)
    const fit = fitIssues(r.records, DEFAULT_PARISH, DEC6, layout)
    expect(fit.filter((i) => i.kind === 'too-long')).toEqual([])
    const doc = await generate(r.records, DEC6)
    expect(doc.labels[0].name).toBe('Maria Guadalupe Esperanza de los Angeles Rodriguez-Fernandez Villanueva')
    expect(Math.max(...doc.labels[0].sizes)).toBeLessThan(28) // shrunk
    expect(Math.max(...doc.labels[1].sizes)).toBe(28) // neighbours unchanged
  })

  it('absurdly long text is flagged as not fitting', () => {
    const rows = base()
    rows[1][7] = 'very long request '.repeat(120)
    const fit = fitIssues(parse(rows).records, DEFAULT_PARISH, DEC6, layout)
    expect(fit[0].kind).toBe('too-long')
  })

  it('extra gift columns are all kept, in order', async () => {
    const rows = base()
    rows[0].push('Fourth gift', 'Gift 5')
    rows[2].push('Puzzle', 'Socks long white') // #302 has three gifts already
    const r = parse(rows)
    expect(r.giftColumns.length).toBe(5)
    expect(r.records[1].gifts.length).toBe(5)
    expect(r.records[1].gifts.slice(3)).toEqual(['Puzzle', 'Socks long white'])
    const doc = await generate(r.records, DEC6)
    expect(doc.labels[1].gifts).toEqual(r.records[1].gifts)
  })

  it('different numbering range, unsorted input', async () => {
    const rows: Row[] = [HEADER, [1002, 'Zed', 'Q', 'F', 7, null, null, 'Ball'], [17, 'Amy', 'R', 'M', 9, null, null, 'Kite'], [5, 'Bo', 'S', null, null, null, null, 'Book']]
    const r = parse(rows)
    expect(r.records.map((x) => x.number)).toEqual(['5', '17', '1002'])
    const doc = await generate(r.records, DEC6)
    expect(doc.labels.map((l) => l.number)).toEqual(['5', '17', '1002'])
    expect(doc.labels[0].gender).toBe('')
    expect(r.issues.some((i) => i.kind === 'missing-gender')).toBe(true)
  })

  it('ignores title/instruction rows above the header, blank rows and notes', () => {
    const rows: Row[] = [
      ['Christmas Angels 2026'],
      ['Please fill one row per child'],
      [],
      HEADER,
      [301, ' Testa ', 'Ejemplo', 'F', 11, null, null, 'Sweater  Black L ', null, null],
      [],
      ['Total: 1'],
    ]
    const r = parse(rows)
    expect(r.issues.filter((i) => i.level === 'error')).toEqual([])
    expect(r.records).toEqual([
      { number: '301', numberValue: 301, name: 'Testa Ejemplo', gender: 'Female', age: '11', gifts: ['Sweater  Black L'], row: 5 },
    ])
  })

  it('accepts a single "Name" column and "#301" style numbers', () => {
    const rows: Row[] = [['Angel No.', 'Name', 'Age', 'Gift 1', 'Gift 2'], ['#12', 'Rosa Diaz', 'adult', 'Coat', 'Hat']]
    const r = parse(rows)
    expect(r.records[0]).toMatchObject({ number: '12', name: 'Rosa Diaz', age: 'adult', gifts: ['Coat', 'Hat'] })
  })
})

describe('layout', () => {
  it('Avery 5168 matches the reference document geometry', async () => {
    const r = parseWorkbookData(workbookBytes(referenceRows()))
    const doc = await generate(r.records, DEC6)
    const xml = doc.documentXml
    expect(xml).toContain('<w:pgSz w:w="12240" w:h="15840"')
    expect(xml).toMatch(/<w:gridCol w:w="5055"\/><w:gridCol w:w="630"\/><w:gridCol w:w="5115"\/>/)
    expect(xml).toMatch(/w:top="720"/)
    expect(xml).toMatch(/w:left="720"/)
    expect(xml).toContain('<w:tblLayout w:type="fixed"/>')
    // 199 labels, 2 per row → 100 rows, all exact 5" and unsplittable; 2 rows per page → 50 pages.
    expect(doc.rowsXml.length).toBe(100)
    for (const row of doc.rowsXml) {
      expect(row).toContain('<w:cantSplit/>')
      expect(row).toMatch(/<w:trHeight w:val="7200" w:hRule="exact"\/>|<w:trHeight w:hRule="exact" w:val="7200"\/>/)
    }
    expect(xml).not.toMatch(/<w:tab\/>|w:type="page"|pageBreakBefore/)
    // No visible borders.
    expect(xml).not.toMatch(/w:val="single"/)
  })

  it('page body fits exactly rowsPerPage rows for every layout (no blank pages, no drift)', () => {
    for (const l of LAYOUTS) {
      const m = pageMargins(l)
      const body = l.pageHeight - m.top - m.bottom
      const rows = l.rowsPerPage * l.labelHeight + (l.rowsPerPage - 1) * l.rowGap
      expect(body - rows, l.id).toBe(BOTTOM_BUFFER)
      expect(body, l.id).toBeLessThan(rows + l.labelHeight) // next row cannot fit
      expect(m.left + m.right + l.columns.reduce((a, b) => a + b, 0), l.id).toBe(l.pageWidth)
    }
  })

  it.each(LAYOUTS.map((l) => l.id))('%s produces every label with the right count per page', async (id) => {
    const r = parseWorkbookData(workbookBytes(referenceRows()))
    const l = getLayout(id)
    const doc = await generate(r.records, DEC6, id)
    expectLabelsMatch(r.records, doc.labels, DEC6)
    const labelRows = Math.ceil(199 / labelsPerPage(l)) * l.rowsPerPage
    const gapRows = l.rowGap ? Math.ceil(199 / labelsPerPage(l)) * (l.rowsPerPage - 1) : 0
    expect(doc.rowsXml.length).toBe(labelRows + gapRows)
    expect(fitIssues(r.records, DEFAULT_PARISH, DEC6, l).filter((i) => i.kind === 'too-long')).toEqual([])
  })
})
