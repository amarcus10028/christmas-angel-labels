import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { parseWorkbookData } from '../src/lib/parse'
import { generateDocxBuffer } from '../src/lib/docx'
import { getLayout } from '../src/lib/layouts'
import { DEFAULT_PARISH } from '../src/lib/types'
import { extractDocx } from './helpers'

// Compares against the parish's own finished document. Both files hold real names, so they
// live in the git-ignored fixtures-private/ folder and this test is skipped when they're absent.
const XLSX = resolve(__dirname, '../fixtures-private/OLL.xlsx')
const DOCX = resolve(__dirname, '../fixtures-private/Christmas_Angel_Labels_2026_December_6.docx')
const have = existsSync(XLSX) && existsSync(DOCX)

describe.skipIf(!have)('matches the parish reference document (December 6)', () => {
  it('every label has identical text, except the reference’s "Age: None" bug', async () => {
    const ref = await extractDocx(new Uint8Array(readFileSync(DOCX)))
    const { records } = parseWorkbookData(new Uint8Array(readFileSync(XLSX)))
    const ours = await extractDocx(
      await generateDocxBuffer({ records, parish: DEFAULT_PARISH, deadlineText: 'Sunday, December 6, 2026', layout: getLayout('avery-5168') }),
    )
    expect(ours.labels.length).toBe(ref.labels.length)
    let fixedAges = 0
    ours.labels.forEach((o, i) => {
      const r = ref.labels[i]
      if (r.age === 'None') {
        fixedAges++
        expect(o.lines).toEqual(r.lines.filter((l) => l !== 'Age: None'))
      } else {
        expect(o.lines).toEqual(r.lines)
        expect(o.sizes).toEqual(r.sizes.map((s) => s || 22).map((s) => (s === 16 ? 22 : s)))
      }
    })
    expect(fixedAges).toBe(5)
  })
})
