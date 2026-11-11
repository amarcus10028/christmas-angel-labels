import type { AngelRecord, ParishSettings } from './types'
import type { LabelLayout } from './layouts'

export type LineStyle = 'title' | 'bold' | 'regular' | 'spacer'

export interface LabelLine {
  text: string
  style: LineStyle
}

export const NO_GIFTS_TEXT = 'Gift choices not provided'

/** Reference type sizes in points, taken from the December 6 document. */
export const BASE_SIZES: Record<LineStyle, number> = {
  title: 14,
  bold: 12,
  regular: 12,
  // Empty paragraphs in the reference render at Word's default 11 pt.
  spacer: 11,
}

export function deadlineLine(deadlineText: string): string {
  return `Return deadline: ${deadlineText}`
}

/** The content of one label, top to bottom, matching the reference document. */
export function labelLines(rec: AngelRecord, parish: ParishSettings, deadlineText: string): LabelLine[] {
  const lines: LabelLine[] = []
  const add = (text: string, style: LineStyle) => lines.push({ text, style })
  const spacer = () => add('', 'spacer')

  add('Be a Christmas Angel to:', 'title')
  spacer()
  add(`ANGEL #${rec.number}`, 'bold')
  spacer()
  add(rec.name, 'bold')
  if (rec.gender) add(rec.gender, 'bold')
  if (rec.age) add(`Age: ${rec.age}`, 'bold')
  spacer()
  add('Gift Choices:', 'bold')
  spacer()
  if (rec.gifts.length) rec.gifts.forEach((g) => add(g, 'bold'))
  else add(NO_GIFTS_TEXT, 'regular')
  spacer()
  const parishName = parish.parishName.trim()
  if (parishName) {
    add(parishName, 'regular')
    spacer()
  }
  parish.returnInstructions
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean)
    .forEach((l) => add(l, 'regular'))
  add(deadlineLine(deadlineText), 'regular')
  return lines
}

// ---------------------------------------------------------------------------
// Text fitting. Word clips text that overflows an exact-height row, so we
// estimate each label's height with Calibri metrics and shrink that label's
// type until it fits.

// Approximate Calibri regular advance widths, in 1/1000 em.
const W: Record<string, number> = {
  ' ': 226, a: 479, b: 525, c: 423, d: 525, e: 498, f: 305, g: 471, h: 525, i: 229, j: 239, k: 455,
  l: 229, m: 799, n: 525, o: 527, p: 525, q: 525, r: 349, s: 391, t: 335, u: 525, v: 452, w: 715,
  x: 433, y: 453, z: 395, A: 579, B: 544, C: 533, D: 615, E: 488, F: 459, G: 631, H: 623, I: 252,
  J: 319, K: 520, L: 420, M: 855, N: 646, O: 662, P: 517, Q: 673, R: 543, S: 459, T: 487, U: 642,
  V: 567, W: 890, X: 519, Y: 487, Z: 468, '.': 252, ',': 250, ':': 268, ';': 268, '-': 306,
  '/': 386, '(': 303, ')': 303, "'": 221, '"': 401, '&': 682, '#': 498, '!': 326, '?': 463,
}
const DIGIT = 507
const OTHER = 600
const BOLD_FACTOR = 1.05
const SAFETY = 1.06
/** Calibri ascent + descent + line gap, as a multiple of the font size. */
export const LINE_HEIGHT = 1.2207

function charWidth(ch: string): number {
  if (ch >= '0' && ch <= '9') return DIGIT
  return W[ch] ?? (/\p{L}/u.test(ch) ? W[ch.normalize('NFD')[0]] ?? OTHER : OTHER)
}

export function textWidthPt(text: string, sizePt: number, bold: boolean): number {
  let w = 0
  for (const ch of text) w += charWidth(ch)
  return (w / 1000) * sizePt * (bold ? BOLD_FACTOR : 1) * SAFETY
}

/** Number of lines Word will need for this text when greedily word-wrapped. */
export function wrappedLineCount(text: string, sizePt: number, bold: boolean, widthPt: number): number {
  if (!text) return 1
  const space = textWidthPt(' ', sizePt, bold)
  let lines = 1
  let cur = 0
  for (const word of text.split(/\s+/).filter(Boolean)) {
    let ww = textWidthPt(word, sizePt, bold)
    if (cur > 0 && cur + space + ww <= widthPt) {
      cur += space + ww
      continue
    }
    if (cur > 0) lines++
    // A single word wider than the line is broken across lines by Word.
    while (ww > widthPt) {
      lines++
      ww -= widthPt
    }
    cur = ww
  }
  return lines
}

export const SCALE_STEPS = [1, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6]

/** Font size in half-points (Word's unit), rounded to a whole half-point. */
export function halfPoints(style: LineStyle, scale: number): number {
  return Math.max(2, Math.round(BASE_SIZES[style] * scale * 2))
}

export function labelHeightPt(lines: LabelLine[], scale: number, widthPt: number): number {
  let h = 0
  for (const l of lines) {
    const size = halfPoints(l.style, scale) / 2
    const n = l.style === 'spacer' ? 1 : wrappedLineCount(l.text, size, l.style !== 'regular', widthPt)
    h += n * size * LINE_HEIGHT
  }
  return h
}

export interface FitResult {
  scale: number
  /** Scale relative to the layout's normal size (1 = no shrinking needed). */
  relative: number
  fits: boolean
}

export function contentBox(layout: LabelLayout, column: number) {
  const width = layout.columns[column] - 2 * 15 - 2 * layout.padSide
  // 6 pt of slack under the text for printer/rounding differences.
  const height = layout.labelHeight - layout.padTop - 120
  return { widthPt: width / 20, heightPt: height / 20 }
}

export function fitLabel(lines: LabelLine[], layout: LabelLayout, column = 0): FitResult {
  // Use the narrowest label column so a label fits wherever it lands.
  const box = layout.columns
    .filter((_, i) => i % 2 === 0)
    .map((_, i) => contentBox(layout, i * 2))
    .reduce((a, b) => (b.widthPt < a.widthPt ? b : a), contentBox(layout, column))
  for (const step of SCALE_STEPS) {
    const scale = layout.baseScale * step
    if (labelHeightPt(lines, scale, box.widthPt) <= box.heightPt) return { scale, relative: step, fits: true }
  }
  const last = SCALE_STEPS[SCALE_STEPS.length - 1]
  return { scale: layout.baseScale * last, relative: last, fits: false }
}
