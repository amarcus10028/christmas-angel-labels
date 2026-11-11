import JSZip from 'jszip'
import * as XLSX from 'xlsx'

export interface ExtractedLabel {
  number: string
  name: string
  gender: string
  age: string
  gifts: string[]
  deadlines: string[]
  lines: string[]
  sizes: number[]
}

export interface ExtractedDoc {
  labels: ExtractedLabel[]
  emptySlots: number
  rowsXml: string[]
  documentXml: string
}

const decode = (s: string) =>
  s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')

/** Pull every label back out of a generated .docx as structured data. */
export async function extractDocx(data: Uint8Array): Promise<ExtractedDoc> {
  const zip = await JSZip.loadAsync(data)
  const xml = await zip.file('word/document.xml')!.async('string')
  const rowsXml = xml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) ?? []
  const labels: ExtractedLabel[] = []
  let emptySlots = 0
  for (const row of rowsXml) {
    const cells = row.match(/<w:tc>[\s\S]*?<\/w:tc>/g) ?? []
    for (const cell of cells) {
      const paras = cell.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) ?? []
      const lines = paras.map((p) => decode((p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join('')))
      const sizes = paras.map((p) => Number(/<w:sz w:val="(\d+)"/.exec(p)?.[1] ?? 0))
      if (!lines.some((l) => l.startsWith('ANGEL #'))) {
        if (cell.includes('w:w="5055"') || cell.includes('w:w="5115"')) emptySlots++
        continue
      }
      labels.push(parseLabel(lines, sizes))
    }
  }
  return { labels, emptySlots, rowsXml, documentXml: xml }
}

function parseLabel(lines: string[], sizes: number[]): ExtractedLabel {
  // Structure: title, '', ANGEL, '', name, [gender], [Age], '', Gift Choices:, '', gifts..., '', parish, '', instructions..., deadline
  const i = lines.findIndex((l) => l.startsWith('ANGEL #'))
  const number = lines[i].slice('ANGEL #'.length)
  const name = lines[i + 2]
  let j = i + 3
  const block: string[] = []
  while (lines[j] !== '') block.push(lines[j++])
  const ageLine = block.find((l) => l.startsWith('Age: ')) ?? ''
  const gender = block.find((l) => !l.startsWith('Age: ')) ?? ''
  const gc = lines.indexOf('Gift Choices:')
  const gifts: string[] = []
  for (let k = gc + 2; lines[k] !== ''; k++) gifts.push(lines[k])
  return {
    number,
    name,
    gender,
    age: ageLine ? ageLine.slice(5) : '',
    gifts,
    deadlines: lines.filter((l) => l.startsWith('Return deadline:')),
    lines,
    sizes,
  }
}

// ---------------------------------------------------------------------------
// Synthetic fixtures with the same shape as the parish's OLL.xlsx (fake names).

const FIRST = ['Ana', 'Ben', 'Carla', 'Diego', 'Elena', 'Felix', 'Gina', 'Hugo', 'Iris', 'Jon', 'Kira', 'Luis', 'Mia', 'Nico', 'Olga', 'Pablo']
const LAST = ['Alvarez', 'Brooks', 'Castro', 'Diaz', 'Ellis', 'Flores', 'Garcia', 'Hale', 'Ibarra', 'Jones']
const GIFTS = ['Sweater Medium', 'Blanket', 'Heater', 'Books/Crayons', 'Boots Black 4/ 4.5', "pajamas women's small any color", 'Pants 10-12', 'Scarf/Gloves']

export type Row = (string | number | null)[]

export const HEADER: Row = ['Angel #', 'first name', 'Last name', 'gender', 'age', null, null, 'first gift', 'second gift', 'Third Gift']

/** 199 recipients, #301–#500 without #429, ages missing for #379–#383, one lowercase gender. */
export function referenceRows(): Row[] {
  const rows: Row[] = [HEADER]
  let k = 0
  for (let n = 301; n <= 500; n++) {
    if (n === 429) continue
    k++
    const gender = n === 350 ? 'm' : k % 2 ? 'M' : 'F'
    const age = n >= 379 && n <= 383 ? null : 4 + (k % 14)
    const giftCount = n % 7 === 0 ? 1 : 3
    const gifts = Array.from({ length: 3 }, (_, g) => (g < giftCount ? GIFTS[(k + g) % GIFTS.length] : null))
    rows.push([n, FIRST[k % FIRST.length], LAST[k % LAST.length], gender, age, null, null, ...gifts])
  }
  return rows
}

export function workbookBytes(rows: Row[], sheetName = 'Sheet1'): Uint8Array {
  const ws = XLSX.utils.aoa_to_sheet(rows)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }))
}
