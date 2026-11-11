import {
  AlignmentType,
  Document,
  HeightRule,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
  BorderStyle,
} from 'docx'
import type { AngelRecord, Issue, ParishSettings } from './types'
import { type LabelLayout, labelsPerPage, pageMargins } from './layouts'
import { fitLabel, halfPoints, labelLines, type LabelLine } from './label'

const FONT = 'Calibri'
const NONE = { style: BorderStyle.NONE, size: 0, color: 'auto' }
const NO_BORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE }

function para(line: LabelLine, scale: number, padSide: number): Paragraph {
  const size = halfPoints(line.style, scale)
  const bold = line.style === 'bold'
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 240 },
    indent: padSide ? { left: padSide, right: padSide } : undefined,
    run: { font: FONT, size, bold },
    children: line.text ? [new TextRun({ text: line.text, font: FONT, size, bold })] : [],
  })
}

function labelCell(lines: LabelLine[] | null, scale: number, width: number, layout: LabelLayout): TableCell {
  const children = lines ? lines.map((l) => para(l, scale, layout.padSide)) : [new Paragraph({ children: [] })]
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: layout.padTop ? { top: layout.padTop, marginUnitType: WidthType.DXA } : undefined,
    children,
  })
}

function gapCell(width: number): TableCell {
  return new TableCell({ width: { size: width, type: WidthType.DXA }, children: [new Paragraph({ children: [] })] })
}

export interface GenerateOptions {
  records: AngelRecord[]
  parish: ParishSettings
  deadlineText: string
  layout: LabelLayout
}

export function buildDocument({ records, parish, deadlineText, layout }: GenerateOptions): Document {
  const cols = layout.columns
  const labelCols = cols.map((_, i) => i).filter((i) => i % 2 === 0)
  const perRow = labelCols.length
  const perPage = labelsPerPage(layout)
  const totalSlots = Math.max(perPage, Math.ceil(records.length / perPage) * perPage)

  const rows: TableRow[] = []
  for (let slot = 0, rowIndex = 0; slot < totalSlots; slot += perRow, rowIndex++) {
    if (layout.rowGap && rowIndex % layout.rowsPerPage !== 0) {
      rows.push(
        new TableRow({
          height: { value: layout.rowGap, rule: HeightRule.EXACT },
          cantSplit: true,
          children: cols.map((w) => gapCell(w)),
        }),
      )
    }
    const cells = cols.map((w, ci) => {
      if (ci % 2 === 1) return gapCell(w)
      const rec = records[slot + ci / 2]
      if (!rec) return labelCell(null, 1, w, layout)
      const lines = labelLines(rec, parish, deadlineText)
      return labelCell(lines, fitLabel(lines, layout).scale, w, layout)
    })
    rows.push(new TableRow({ height: { value: layout.labelHeight, rule: HeightRule.EXACT }, cantSplit: true, children: cells }))
  }

  const table = new Table({
    rows,
    columnWidths: cols,
    width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    layout: TableLayoutType.FIXED,
    indent: { size: -15, type: WidthType.DXA },
    margins: { left: 15, right: 15, top: 0, bottom: 0 },
    borders: NO_BORDERS,
  })

  const m = pageMargins(layout)
  return new Document({
    creator: 'Christmas Angel Label Maker',
    title: 'Christmas Angel Labels',
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 22 },
          paragraph: { spacing: { before: 0, after: 0, line: 240 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: layout.pageWidth, height: layout.pageHeight },
            margin: { ...m, header: 0, footer: 0, gutter: 0 },
          },
        },
        // Word requires a paragraph after a table; keep it 1 pt so it can never start a blank page.
        children: [table, new Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: 'exact' }, run: { size: 2 }, children: [] })],
      },
    ],
  })
}

export async function generateDocx(opts: GenerateOptions): Promise<Blob> {
  return Packer.toBlob(buildDocument(opts))
}

export async function generateDocxBuffer(opts: GenerateOptions): Promise<Uint8Array> {
  const buf = await Packer.toBuffer(buildDocument(opts))
  return new Uint8Array(buf)
}

/** Labels whose text is long enough that it must be shrunk (or won't fit at all). */
export function fitIssues(records: AngelRecord[], parish: ParishSettings, deadlineText: string, layout: LabelLayout): Issue[] {
  const issues: Issue[] = []
  for (const rec of records) {
    const f = fitLabel(labelLines(rec, parish, deadlineText), layout)
    const who = `Angel #${rec.number} (${rec.name})`
    if (!f.fits)
      issues.push({
        level: 'warning',
        kind: 'too-long',
        row: rec.row,
        message: `Row ${rec.row}: ${who} has too much text to fit on one label, even in smaller type. Please shorten the name or gift choices in the spreadsheet.`,
      })
    // Small reductions aren't worth a volunteer's attention; mention only noticeable ones.
    else if (f.relative < 0.9)
      issues.push({
        level: 'warning',
        kind: 'long-text',
        row: rec.row,
        message: `Row ${rec.row}: ${who} has long gift descriptions, so this label’s type will be ${Math.round(f.relative * 100)}% of normal size to fit.`,
      })
  }
  return issues
}
