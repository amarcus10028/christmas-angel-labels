export interface AngelRecord {
  /** Angel number as typed (digits only, no "#"). */
  number: string
  /** Numeric value used for sorting and duplicate checks. */
  numberValue: number
  name: string
  /** Display text, e.g. "Male". Empty when not provided. */
  gender: string
  /** Display text, e.g. "12" or "adult". Empty when not provided. */
  age: string
  gifts: string[]
  /** 1-based Excel row number. */
  row: number
}

export type IssueLevel = 'error' | 'warning' | 'info'

export type IssueKind =
  | 'unreadable'
  | 'no-header'
  | 'no-records'
  | 'duplicate'
  | 'missing-number'
  | 'missing-name'
  | 'bad-number'
  | 'missing-age'
  | 'missing-gender'
  | 'no-gifts'
  | 'long-text'
  | 'too-long'
  | 'gap'
  | 'note'

export interface Issue {
  level: IssueLevel
  kind: IssueKind
  message: string
  row?: number
}

export interface ParseResult {
  records: AngelRecord[]
  issues: Issue[]
  sheetName?: string
  /** Header text for each gift column found, in order. */
  giftColumns: string[]
}

export interface ParishSettings {
  parishName: string
  /** Each line is printed as its own centered line on the label. */
  returnInstructions: string
}

export const DEFAULT_PARISH: ParishSettings = {
  parishName: 'Church of St. Ignatius Loyola',
  returnInstructions: 'Please return your wrapped gift to the\nParish House at 980 Park Avenue.',
}
