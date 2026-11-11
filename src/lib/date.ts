const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]

/** Parse "YYYY-MM-DD" (the value of an <input type="date">). Returns null if invalid. */
export function parseIsoDate(value: string): { y: number; m: number; d: number } | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value.trim())
  if (!m) return null
  const y = +m[1], mo = +m[2], d = +m[3]
  const dt = new Date(Date.UTC(y, mo - 1, d))
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null
  return { y, m: mo, d }
}

/** "2026-12-06" → "Sunday, December 6, 2026". Computed in UTC so the time zone can't shift the day. */
export function formatDeadline(value: string): string | null {
  const p = parseIsoDate(value)
  if (!p) return null
  const wd = new Date(Date.UTC(p.y, p.m - 1, p.d)).getUTCDay()
  return `${WEEKDAYS[wd]}, ${MONTHS[p.m - 1]} ${p.d}, ${p.y}`
}

export function outputFilename(value: string): string {
  const p = parseIsoDate(value)
  return `Christmas_Angel_Labels_${p ? p.y : new Date().getFullYear()}.docx`
}
