import * as XLSX from 'xlsx'

export const TEMPLATE_FILENAME = 'Christmas_Angel_Spreadsheet_Template.xlsx'

export const TEMPLATE_HEADERS = ['Angel #', 'first name', 'Last name', 'gender', 'age', 'first gift', 'second gift', 'Third Gift']

const TIPS = [
  'How to fill in the Angels sheet',
  '',
  '• One row per angel. Keep the headings in row 1.',
  '• Required for every angel: the angel number and the name. Each angel number must be different.',
  '• Optional: gender (M or F), age, and the gift choices. Anything left blank is simply left off the label.',
  '• Gifts print exactly as typed, so include sizes and colors (for example: Sweater Pink Medium).',
  '• Need more gifts? Add another column after Third Gift with “gift” in the heading, such as “fourth gift”.',
  '• Skipped numbers are fine, and the rows don’t need to be in order.',
  '',
  'When you’re done, save the file and drop it into the Christmas Angel Label Maker.',
]

/** A blank workbook with the expected headings, plus a second sheet of instructions. */
export function buildTemplate(): Uint8Array {
  const wb = XLSX.utils.book_new()

  // Headings only: an example row could end up printed as a real label if it isn't deleted.
  const angels = XLSX.utils.aoa_to_sheet([TEMPLATE_HEADERS])
  angels['!cols'] = [10, 16, 18, 9, 7, 30, 30, 30].map((wch) => ({ wch }))
  XLSX.utils.book_append_sheet(wb, angels, 'Angels')

  const tips = XLSX.utils.aoa_to_sheet(TIPS.map((t) => [t]))
  tips['!cols'] = [{ wch: 110 }]
  XLSX.utils.book_append_sheet(wb, tips, 'Instructions')

  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }))
}
