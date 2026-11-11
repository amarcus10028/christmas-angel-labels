// Writes sample .docx files for visual checking: OUT=<dir> npx vitest run scripts/sample-docs.test.ts
import { it } from 'vitest'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { parseWorkbookData } from '../src/lib/parse'
import { generateDocxBuffer } from '../src/lib/docx'
import { LAYOUTS } from '../src/lib/layouts'
import { DEFAULT_PARISH } from '../src/lib/types'
import { referenceRows, workbookBytes } from '../tests/helpers'

const OUT = process.env.OUT
it.skipIf(!OUT)('write samples', async () => {
  mkdirSync(OUT!, { recursive: true })
  const real = resolve(__dirname, '../fixtures-private/OLL.xlsx')
  const src = existsSync(real) ? new Uint8Array(readFileSync(real)) : workbookBytes(referenceRows())
  const { records } = parseWorkbookData(src)
  const deadlineText = 'Sunday, December 6, 2026'
  for (const layout of LAYOUTS) {
    writeFileSync(`${OUT}/${layout.id}.docx`, await generateDocxBuffer({ records, parish: DEFAULT_PARISH, deadlineText, layout }))
  }
  // Single-page excerpts (first, middle, last page) for renderers that only draw page 1.
  const per = 4
  const pages = Math.ceil(records.length / per)
  for (const [tag, page] of [['first', 0], ['middle', Math.floor(pages / 2)], ['last', pages - 1]] as const) {
    const recs = records.slice(page * per, page * per + per)
    writeFileSync(`${OUT}/page-${tag}.docx`, await generateDocxBuffer({ records: recs, parish: DEFAULT_PARISH, deadlineText, layout: LAYOUTS[0] }))
  }
  const rows = referenceRows()
  rows[1][1] = 'Maria Guadalupe Esperanza de los Angeles'
  rows[1][2] = 'Rodriguez-Fernandez Villanueva'
  rows[1][7] = 'Brandy Melville zip-up hoodie, medium, navy or green or gray, no logos please, women’s fit'
  rows[1][8] = 'Winter coat with hood, black, women’s size small petite, warm lining, long length below knee'
  rows[1][9] = 'Art supplies: watercolor paints, brushes, sketchbook (spiral, 9×12), colored pencils 24-pack'
  rows[2].push('Fourth gift: puzzle 500 pieces', 'Fifth gift: socks')
  rows[0].push('Gift 4', 'Gift 5')
  rows[3][7] = null
  const stress = parseWorkbookData(workbookBytes(rows.slice(0, 8))).records
  for (const layout of LAYOUTS) {
    writeFileSync(`${OUT}/stress-${layout.id}.docx`, await generateDocxBuffer({ records: stress, parish: DEFAULT_PARISH, deadlineText: 'Wednesday, September 30, 2026', layout }))
  }
})
