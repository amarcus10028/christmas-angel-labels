import { useEffect, useMemo, useRef, useState } from 'react'
import './App.css'
import churchPhoto from './assets/church.jpg'
import { parseWorkbookData } from './lib/parse'
import { fitIssues, generateDocx } from './lib/docx'
import { formatDeadline, outputFilename } from './lib/date'
import { DEFAULT_LAYOUT_ID, getLayout, LAYOUTS } from './lib/layouts'
import { DEFAULT_PARISH, type Issue, type IssueKind, type ParishSettings, type ParseResult } from './lib/types'

const SETTINGS_KEY = 'angel-labels:settings:v1'

interface Saved {
  parish: ParishSettings
  layoutId: string
}

function loadSaved(): Saved {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY)
    if (raw) {
      const s = JSON.parse(raw) as Partial<Saved>
      return {
        parish: { ...DEFAULT_PARISH, ...(s.parish ?? {}) },
        layoutId: getLayout(s.layoutId).id,
      }
    }
  } catch {
    /* storage unavailable — use defaults */
  }
  return { parish: DEFAULT_PARISH, layoutId: DEFAULT_LAYOUT_ID }
}

function save(s: Saved) {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(s))
  } catch {
    /* ignore */
  }
}

const WARNING_GROUPS: { kind: IssueKind; title: string }[] = [
  { kind: 'too-long', title: 'Too much text to fit' },
  { kind: 'long-text', title: 'Long text (printed a little smaller)' },
  { kind: 'no-gifts', title: 'No gift choices' },
  { kind: 'missing-age', title: 'No age' },
  { kind: 'missing-gender', title: 'No gender' },
]

function plural(n: number, one: string, many = one + 's') {
  return `${n} ${n === 1 ? one : many}`
}

export default function App() {
  const [initial] = useState(loadSaved)
  const [fileName, setFileName] = useState<string | null>(null)
  const [result, setResult] = useState<ParseResult | null>(null)
  const [reading, setReading] = useState(false)
  const [deadline, setDeadline] = useState('')
  const [parish, setParish] = useState<ParishSettings>(initial.parish)
  const [layoutId, setLayoutId] = useState(initial.layoutId)
  const [dragOver, setDragOver] = useState(false)
  const [busy, setBusy] = useState(false)
  const [done, setDone] = useState<{ name: string; key: string } | null>(null)
  const [uploadId, setUploadId] = useState(0)
  const [failure, setFailure] = useState<string | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => save({ parish, layoutId }), [parish, layoutId])

  const layout = getLayout(layoutId)
  // The "downloaded" message only shows while nothing has changed since that download.
  const inputsKey = JSON.stringify([uploadId, deadline, layoutId, parish])
  const doneName = done && done.key === inputsKey ? done.name : null
  const deadlineText = formatDeadline(deadline)

  const issues: Issue[] = useMemo(() => {
    if (!result) return []
    const sample = deadlineText ?? 'Wednesday, September 30, 2026'
    return [...result.issues, ...fitIssues(result.records, parish, sample, layout)]
  }, [result, parish, deadlineText, layout])

  const errors = issues.filter((i) => i.level === 'error')
  const warnings = issues.filter((i) => i.level === 'warning')
  const infos = issues.filter((i) => i.level === 'info')
  const count = (k: IssueKind) => issues.filter((i) => i.kind === k).length

  async function handleFile(file: File | undefined) {
    if (!file) return
    setFileName(file.name)
    setFailure(null)
    setReading(true)
    try {
      const buf = await file.arrayBuffer()
      setResult(parseWorkbookData(new Uint8Array(buf)))
      setUploadId((n) => n + 1)
    } catch {
      setResult({
        records: [],
        giftColumns: [],
        issues: [{ level: 'error', kind: 'unreadable', message: 'This file could not be read. Please try choosing it again.' }],
      })
    } finally {
      setReading(false)
      if (inputRef.current) inputRef.current.value = ''
    }
  }

  async function handleGenerate() {
    if (!result || !deadlineText || errors.length) return
    setBusy(true)
    setFailure(null)
    try {
      const blob = await generateDocx({ records: result.records, parish, deadlineText, layout })
      const name = outputFilename(deadline)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = name
      document.body.appendChild(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(url), 10_000)
      setDone({ name, key: inputsKey })
    } catch {
      setFailure('Something went wrong while making the document. Please try again.')
    } finally {
      setBusy(false)
    }
  }

  const blocker = !result
    ? 'Upload your spreadsheet first.'
    : errors.length
      ? 'Fix the spreadsheet problems listed above, save it, and upload it again.'
      : !deadlineText
        ? 'Choose a gift return deadline.'
        : null

  const parishChanged = parish.parishName !== DEFAULT_PARISH.parishName || parish.returnInstructions !== DEFAULT_PARISH.returnInstructions

  return (
    <div className="page">
      <figure className="hero">
        <img
          src={churchPhoto}
          alt="The Church of St. Ignatius Loyola and Loyola School on Park Avenue, New York City"
          width={1440}
          height={810}
        />
        <figcaption>
          Photo:{' '}
          <a href="https://commons.wikimedia.org/wiki/File:Ignatius_Loyola_complex_2021_jeh.jpg" target="_blank" rel="noopener noreferrer">
            Jim.henderson
          </a>
          ,{' '}
          <a href="https://creativecommons.org/licenses/by/4.0/" target="_blank" rel="noopener noreferrer">CC BY 4.0</a>, via Wikimedia
          Commons (resized)
        </figcaption>
      </figure>
      <header className="masthead">
        <div className="star" aria-hidden="true">✦</div>
        <h1>Christmas Angel Label Maker</h1>
        <p>Turn your Christmas Angel spreadsheet into ready-to-print gift labels.</p>
      </header>

      <main>
        <section className="card" aria-labelledby="s-upload">
          <h2 id="s-upload"><span className="step">1</span> Upload Spreadsheet</h2>
          <details className="howto">
            <summary>How to set up your spreadsheet</summary>
            <p>
              Use one row per angel, with a row of column headings at the top. Name the columns like this
              (capital letters don’t matter, and the columns can be in any order):
            </p>
            <div className="sample" role="region" aria-label="Example spreadsheet">
              <table>
                <thead>
                  <tr>
                    <th>Angel #</th><th>first name</th><th>Last name</th><th>gender</th><th>age</th>
                    <th>first gift</th><th>second gift</th><th>Third Gift</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td>301</td><td>Jane</td><td>Example</td><td>F</td><td>9</td>
                    <td>Sweater Pink Medium</td><td>Books</td><td>Heater</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <ul>
              <li><strong>Angel #</strong> and the <strong>name</strong> are required. Every angel number must be different.</li>
              <li><strong>gender</strong> (M or F), <strong>age</strong> and gifts are optional. If they’re blank, that line is left off the label.</li>
              <li>Need more gifts? Add another column with “gift” in the heading, such as <strong>fourth gift</strong>.</li>
              <li>Gifts print exactly as typed, so include sizes and colors.</li>
              <li>Skipped numbers are fine. Rows don’t need to be in order.</li>
            </ul>
          </details>
          <label
            className={`drop ${dragOver ? 'over' : ''} ${fileName ? 'has-file' : ''}`}
            onDragOver={(e) => {
              e.preventDefault()
              setDragOver(true)
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault()
              setDragOver(false)
              handleFile(e.dataTransfer.files?.[0])
            }}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              onChange={(e) => handleFile(e.target.files?.[0])}
            />
            {fileName ? (
              <>
                <span className="drop-file">{reading ? 'Reading…' : <>✓ Using <strong>{fileName}</strong></>}</span>
                <span className="drop-change">Wrong file? Click here to choose a different spreadsheet</span>
              </>
            ) : (
              <>
                <span className="drop-title">Drop your Excel spreadsheet here</span>
                <span className="drop-hint">or click here to choose the file (.xlsx)</span>
              </>
            )}
          </label>
          <p className="privacy">Your spreadsheet stays on this computer. Nothing is uploaded to the internet.</p>
        </section>

        <section className="card" aria-labelledby="s-date">
          <h2 id="s-date"><span className="step">2</span> Return Deadline</h2>
          <label className="field">
            <span>Gift Return Deadline</span>
            <input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} />
          </label>
          {deadlineText ? (
            <p className="preview">Labels will say: <strong>Return deadline: {deadlineText}</strong></p>
          ) : (
            <p className="muted">Choose the last day gifts can be returned.</p>
          )}
          <label className="field compact">
            <span>Label sheets</span>
            <select value={layoutId} onChange={(e) => setLayoutId(e.target.value)}>
              {LAYOUTS.map((l) => (
                <option key={l.id} value={l.id}>
                  {l.name} — {l.description}
                  {l.id === DEFAULT_LAYOUT_ID ? ' (usual)' : ''}
                </option>
              ))}
            </select>
          </label>
        </section>

        <section className="card" aria-labelledby="s-check" aria-live="polite">
          <h2 id="s-check"><span className="step">3</span> Spreadsheet Check</h2>
          {!result ? (
            <p className="muted">Upload a spreadsheet to check it.</p>
          ) : (
            <>
              <p className={`summary ${errors.length ? 'bad' : 'good'}`}>
                {errors.length ? '⚠️' : '✅'} {plural(result.records.length, 'angel')} found · {plural(count('duplicate'), 'duplicate')} ·{' '}
                {plural(count('missing-name') + count('missing-number') + count('bad-number'), 'missing name or number', 'missing names or numbers')}
                {count('missing-age') ? ` · ${plural(count('missing-age'), 'missing age')}` : ''}
                {warnings.length ? ` · ${plural(warnings.length, 'note')} to review` : ''}
              </p>

              {errors.length > 0 && (
                <div className="errors" role="alert">
                  <p><strong>Please fix {errors.length === 1 ? 'this' : 'these'} in Excel, save, and upload the file again:</strong></p>
                  <ul>
                    {errors.slice(0, 50).map((e, i) => <li key={i}>{e.message}</li>)}
                  </ul>
                  {errors.length > 50 && <p>…and {errors.length - 50} more.</p>}
                </div>
              )}

              {(warnings.length > 0 || infos.length > 0) && (
                <details className="issues">
                  <summary>Review Spreadsheet Issues</summary>
                  <p className="muted">These won’t stop the labels from being made.</p>
                  {WARNING_GROUPS.map(({ kind, title }) => {
                    const list = warnings.filter((w) => w.kind === kind)
                    if (!list.length) return null
                    return (
                      <div key={kind} className="group">
                        <h3>{title} ({list.length})</h3>
                        <ul>{list.slice(0, 40).map((w, i) => <li key={i}>{w.message}</li>)}</ul>
                        {list.length > 40 && <p className="muted">…and {list.length - 40} more.</p>}
                      </div>
                    )
                  })}
                  {infos.map((i, k) => <p key={k} className="info">ℹ️ {i.message}</p>)}
                </details>
              )}
            </>
          )}
        </section>

        <section className="card generate" aria-labelledby="s-gen">
          <h2 id="s-gen"><span className="step">4</span> Generate Document</h2>
          <button type="button" className="primary" disabled={!!blocker || busy} onClick={handleGenerate}>
            {busy ? 'Making your document…' : 'Generate Word Document'}
          </button>
          {blocker && <p className="muted">{blocker}</p>}
          {doneName && (
            <p className="done">
              ✅ <strong>{doneName}</strong> has been downloaded — check your Downloads folder. You can change the date and generate again at any time.
            </p>
          )}
          {failure && <p className="errors">{failure}</p>}
        </section>

        <details className="card settings">
          <summary>Parish Settings</summary>
          <p className="muted">Most people never need to change these. Changes are remembered on this computer.</p>
          <label className="field">
            <span>Parish name</span>
            <input type="text" value={parish.parishName} onChange={(e) => setParish({ ...parish, parishName: e.target.value })} />
          </label>
          <label className="field">
            <span>Return instructions (each line prints on its own line)</span>
            <textarea rows={3} value={parish.returnInstructions} onChange={(e) => setParish({ ...parish, returnInstructions: e.target.value })} />
          </label>
          {parishChanged && (
            <button type="button" className="link" onClick={() => setParish(DEFAULT_PARISH)}>Restore the original wording</button>
          )}
        </details>
      </main>

      <footer>
        Made by{' '}
        <a href="https://www.linkedin.com/in/alexmarcus1/" target="_blank" rel="noopener noreferrer">
          Alex Marcus
        </a>{' '}
        for the Christmas Angel ministry at the Church of St. Ignatius Loyola · Works entirely in your browser
      </footer>
    </div>
  )
}
