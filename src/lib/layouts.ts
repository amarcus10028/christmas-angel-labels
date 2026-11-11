/**
 * Label sheet geometry. All lengths are in twips (1/1440 inch).
 *
 * Labels are laid out as one borderless, fixed-layout table. Every label row has an
 * exact height and cannot split, so each page holds exactly `rowsPerPage` rows and
 * Word starts the next row at the top margin of the next page. No spaces, tabs or
 * manual page breaks are used for positioning.
 */
export interface LabelLayout {
  id: string
  name: string
  description: string
  pageWidth: number
  pageHeight: number
  marginTop: number
  marginLeft: number
  /** Grid columns left to right. Label columns are at even indexes, gaps at odd. */
  columns: number[]
  rowsPerPage: number
  labelHeight: number
  /** Vertical gap between label rows on the same page. */
  rowGap: number
  /** Extra padding inside each label. */
  padTop: number
  padSide: number
  /** Multiplier applied to the reference type sizes (14 / 12 pt). */
  baseScale: number
}

const IN = 1440

export const LAYOUTS: LabelLayout[] = [
  {
    // Geometry copied from the parish's existing December 6 document (Avery 5168 template).
    id: 'avery-5168',
    name: 'Avery 5168 / 8168',
    description: '3½″ × 5″ · 4 labels per sheet',
    pageWidth: 12240,
    pageHeight: 15840,
    marginTop: 720,
    marginLeft: 720,
    columns: [5055, 630, 5115],
    rowsPerPage: 2,
    labelHeight: 7200,
    rowGap: 0,
    padTop: 0,
    padSide: 126,
    baseScale: 1,
  },
  {
    id: 'avery-5164',
    name: 'Avery 5164 / 8164',
    description: '3⅓″ × 4″ · 6 labels per sheet',
    pageWidth: 12240,
    pageHeight: 15840,
    marginTop: 720,
    marginLeft: 225,
    columns: [5760, 270, 5760],
    rowsPerPage: 3,
    labelHeight: 4800,
    rowGap: 0,
    padTop: 72,
    padSide: 144,
    baseScale: 0.8,
  },
  {
    id: 'avery-5126',
    name: 'Avery 5126 / 8126',
    description: '5½″ × 8½″ half sheet · 2 labels per sheet',
    pageWidth: 12240,
    pageHeight: 15840,
    marginTop: Math.round(0.4 * IN),
    marginLeft: 720,
    columns: [10800],
    rowsPerPage: 2,
    labelHeight: Math.round(4.7 * IN),
    rowGap: Math.round(0.8 * IN),
    padTop: 0,
    padSide: 0,
    baseScale: 1.15,
  },
  {
    id: 'avery-5165',
    name: 'Avery 5165 / 8165',
    description: '8½″ × 11″ full sheet · 1 label per sheet',
    pageWidth: 12240,
    pageHeight: 15840,
    marginTop: 720,
    marginLeft: 720,
    columns: [10800],
    rowsPerPage: 1,
    labelHeight: 10 * IN,
    rowGap: 0,
    padTop: 360,
    padSide: 0,
    baseScale: 1.75,
  },
]

export const DEFAULT_LAYOUT_ID = 'avery-5168'

export function getLayout(id: string | undefined): LabelLayout {
  return LAYOUTS.find((l) => l.id === id) ?? LAYOUTS[0]
}

export function labelsPerPage(layout: LabelLayout): number {
  return Math.ceil(layout.columns.length / 2) * layout.rowsPerPage
}

/** Small buffer below the last row so Word's required final paragraph never makes a blank page. */
export const BOTTOM_BUFFER = 144

export function pageMargins(layout: LabelLayout) {
  const used = layout.columns.reduce((a, b) => a + b, 0)
  const body = layout.rowsPerPage * layout.labelHeight + (layout.rowsPerPage - 1) * layout.rowGap + BOTTOM_BUFFER
  return {
    top: layout.marginTop,
    left: layout.marginLeft,
    right: Math.max(0, layout.pageWidth - layout.marginLeft - used),
    bottom: Math.max(0, layout.pageHeight - layout.marginTop - body),
  }
}
