import { useRef, useState } from 'react'
import { cleanRows, FIELDS, sheetData, suggestMapping } from './cleaner'
import './App.css'

const initialOptions = { contact: 'phone', country: 'US', requireName: true, splitNames: true, multi: 'valid', dedupe: 'phone', keep: 'complete', extra: [] }
const PAGE_SIZE = 20

export default function App() {
  const input = useRef(null), request = useRef(0)
  const [stage, setStage] = useState(0)
  const [fileName, setFileName] = useState('')
  const [sheets, setSheets] = useState({})
  const [sheet, setSheet] = useState('')
  const [headerRow, setHeaderRow] = useState(0)
  const [data, setData] = useState({ headers: [], rows: [] })
  const [mapping, setMapping] = useState({})
  const [options, setOptions] = useState(initialOptions)
  const [results, setResults] = useState([])
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [query, setQuery] = useState('')
  const [filter, setFilter] = useState('all')
  const [page, setPage] = useState(0)
  const [details, setDetails] = useState(null)
  const [drag, setDrag] = useState(false)
  const option = (key, value) => setOptions(prev => ({ ...prev, [key]: value }))
  function configure(matrix, row = 0) {
    const next = sheetData(matrix, row)
    const suggested = suggestMapping(next.headers)
    setData(next); setMapping(suggested); setHeaderRow(row)
    setOptions(prev => ({ ...prev, extra: next.headers.map((_, i) => i).filter(i => !Object.values(suggested).includes(String(i))) }))
  }
  async function load(file) {
    if (!file) return
    const token = ++request.current
    setBusy(true); setError(''); setFileName(''); setSheets({}); setData({ headers: [], rows: [] }); setResults([]); setStage(0)
    try {
      if (!/\.(csv|xlsx|xls)$/i.test(file.name)) throw new Error('Choose a CSV, XLSX, or XLS file.')
      if (file.size > 25 * 1024 * 1024) throw new Error('This file exceeds 25 MB. Split it into smaller files and try again.')
      const XLSX = await import('xlsx')
      const buffer = await file.arrayBuffer()
      if (token !== request.current) return
      const workbook = XLSX.read(buffer, { type: 'array', cellText: true })
      const parsed = Object.fromEntries(workbook.SheetNames.map(name => [name, XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false, blankrows: true })]))
      const first = workbook.SheetNames.find(name => parsed[name].some(row => row.some(v => String(v).trim())))
      if (!first) throw new Error('This file is empty. Add a header row and at least one contact.')
      setSheets(parsed); setSheet(first); configure(parsed[first]); setFileName(file.name); setStage(1)
    } catch (e) { if (token === request.current) setError(e.message || 'Could not read this file. Try exporting it as CSV.') }
    finally { if (token === request.current) setBusy(false); if (input.current) input.current.value = '' }
  }
  const used = Object.values(mapping).filter(v => v !== '-1')
  const mappingError = new Set(used).size !== used.length ? 'Map each source column to only one field.' : options.requireName && mapping.first_name === '-1' && mapping.full_name === '-1' ? 'Map a first name or full name column, or turn off the name requirement.' : options.contact === 'phone' && mapping.phone === '-1' ? 'Map a phone column for this contact rule.' : options.contact === 'email' && mapping.email === '-1' ? 'Map an email column for this contact rule.' : options.contact === 'either' && mapping.phone === '-1' && mapping.email === '-1' ? 'Map a phone or email column.' : ''
  async function run() {
    setBusy(true); setError('')
    await new Promise(resolve => setTimeout(resolve, 30))
    try { setResults(cleanRows(data.rows, data.headers, mapping, options)); setStage(2); setPage(0); setQuery(''); setFilter('all'); setDetails(null) }
    catch { setError('Cleaning failed. Check the selected worksheet and columns, then try again.') }
    finally { setBusy(false) }
  }
  const kept = results.filter(r => !r.reason || r.restored)
  const removed = results.filter(r => r.reason && !r.restored)
  const filtered = results.filter(r => (filter === 'all' || (filter === 'kept' ? !r.reason || r.restored : filter === 'removed' ? r.reason && !r.restored : r.warnings.length > 0)) && `${r.id} ${r.reason} ${r.warnings.join(' ')} ${r.cells.join(' ')} ${Object.values(r.output).join(' ')}`.toLowerCase().includes(query.toLowerCase()))
  const pages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE))
  const currentPage = Math.min(page, pages - 1)
  async function download(kind, format = 'csv') {
    setBusy(true); setError('')
    try {
      const XLSX = await import('xlsx')
      const source = kind === 'cleaned' ? kept : kind === 'removed' ? removed : results
      const headers = kind === 'cleaned' ? Object.keys(source[0]?.output || {}) : ['Source row', 'Status', 'Reason', 'Warnings', ...data.headers.map((h, i) => `${h} [${i + 1}]`)]
      const rows = source.map(r => kind === 'cleaned' ? headers.map(h => r.output[h]) : [r.id, r.restored ? 'Restored' : r.reason ? 'Removed' : 'Kept', r.reason, r.warnings.join('; '), ...r.cells])
      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows])
      const base = fileName.replace(/\.[^.]+$/, '')
      if (format === 'xlsx') { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Contacts'); XLSX.writeFile(wb, `${base}_${kind}.xlsx`) }
      else {
        // Escape spreadsheet formula prefixes in CSV; XLSX exports preserve string cells.
        const safe = XLSX.utils.aoa_to_sheet([headers, ...rows].map(row => row.map(v => typeof v === 'string' && /^[\s]*[=+@-]/.test(v) ? `'${v}` : v)))
        const url = URL.createObjectURL(new Blob(['\ufeff', XLSX.utils.sheet_to_csv(safe)], { type: 'text/csv;charset=utf-8' }))
        const a = document.createElement('a'); a.href = url; a.download = `${base}_${kind}.csv`; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000)
      }
      setStage(3)
    } catch { setError('Export failed. Please try again.') }
    finally { setBusy(false) }
  }
  function reset() { request.current++; setStage(0); setFileName(''); setSheets({}); setData({ headers: [], rows: [] }); setResults([]); setError(''); setDetails(null); setBusy(false) }
  return <div className="app">
    <header className="header"><div className="header-inner"><span className="logo-mark">◈</span><div><h1>Contact List Cleaner</h1><span className="subtitle">for Relentless Digital</span></div><span className="privacy">Files stay in your browser</span></div></header>
    <main className="main">
      <nav aria-label="Progress" className="steps">{['Upload', 'Configure', 'Review', 'Export'].map((label, i) => <span key={label} aria-current={stage === i ? 'step' : undefined} className={stage >= i ? 'active' : ''}><b>{i + 1}</b>{label}</span>)}</nav>
      {error && <div role="alert" className="error">{error}</div>}
      {busy && <p role="status" className="notice">Processing your file…</p>}
      <input ref={input} type="file" accept=".csv,.xlsx,.xls" className="sr-only" tabIndex={-1} onChange={e => load(e.target.files[0])} />
      {stage === 0 ? <section className="card upload-card"><span className="eyebrow">A cleaner list. A clearer next step.</span><h2>Get your contacts ready to use</h2><p className="hint">Map your columns, choose your rules, and review every change before exporting.</p><button disabled={busy} className={`dropzone ${drag ? 'dragover' : ''}`} onClick={() => input.current.click()} onDragOver={e => { e.preventDefault(); setDrag(true) }} onDragLeave={() => setDrag(false)} onDrop={e => { e.preventDefault(); setDrag(false); if (!busy) load(e.dataTransfer.files[0]) }}><span className="drop-icon">↑</span><strong>Drop your contact list here</strong><span>or browse files · CSV, XLSX, XLS · up to 25 MB</span></button><div className="intro-grid"><div><h3>Keep the context</h3><p>Notes and extra columns stay intact.</p></div><div><h3>Stay in control</h3><p>Review removals and restore contacts.</p></div><div><h3>Export with confidence</h3><p>Download your contacts and audit trail.</p></div></div></section> : <>
      <div className="file-bar"><div><strong>{fileName}</strong><span>{data.rows.length.toLocaleString()} contacts · {data.headers.length} columns</span></div><button className="btn btn-ghost" disabled={busy} onClick={reset}>New file</button></div>
      {stage === 1 ? <>
        <section className="card"><h2>Confirm your columns</h2><p className="hint">We suggest exact header matches. Check the examples and adjust any field before cleaning.</p><div className="two-col"><label>Worksheet<select value={sheet} onChange={e => { setSheet(e.target.value); configure(sheets[e.target.value]) }}>{Object.keys(sheets).map(name => <option key={name}>{name}</option>)}</select></label><label>Header row<input type="number" min="1" max={Math.max(1, sheets[sheet]?.length || 1)} value={headerRow + 1} onChange={e => configure(sheets[sheet], Math.max(0, Math.min((sheets[sheet]?.length || 1) - 1, Number(e.target.value) - 1)))} /></label></div><div className="mapping-grid">{Object.entries(FIELDS).map(([key, label]) => <label key={key}>{label}<select value={mapping[key] ?? '-1'} onChange={e => setMapping(prev => ({ ...prev, [key]: e.target.value }))}><option value="-1">Not mapped</option>{data.headers.map((h, i) => <option value={i} key={i}>{i + 1}. {h}</option>)}</select><small>{data.rows.slice(0, 2).map(r => r.cells[Number(mapping[key])]).filter(Boolean).join(' · ') || 'No sample value'}</small></label>)}</div>{!data.rows.length && <p className="error">No contacts below this header. Select another worksheet or header row.</p>}</section>
        <section className="card"><h2>Choose your cleaning rules</h2><p className="hint">Invalid values are flagged. Validation checks format and numbering rules, not whether a contact is reachable.</p><div className="two-col"><label>Contact requirement<select value={options.contact} onChange={e => option('contact', e.target.value)}><option value="phone">Valid phone required</option><option value="either">Allow email-only contacts</option><option value="email">Valid email required</option></select></label><label>Default phone country<select value={options.country} onChange={e => option('country', e.target.value)}><option value="US">United States</option><option value="CA">Canada</option><option value="GB">United Kingdom</option><option value="AU">Australia</option></select><small>International numbers with + use their own country code.</small></label><label>Find duplicates by<select value={options.dedupe} onChange={e => option('dedupe', e.target.value)}><option value="phone">Phone and extension</option><option value="email">Email</option><option value="either">Phone or email</option><option value="none">Keep all (shared numbers allowed)</option></select></label><label>When duplicates match<select value={options.keep} onChange={e => option('keep', e.target.value)}><option value="complete">Keep the most complete contact</option><option value="first">Keep the first contact</option></select></label><label>Multiple phones or emails<select value={options.multi} onChange={e => option('multi', e.target.value)}><option value="valid">Use the first valid value</option><option value="first">Use only the first value</option></select><small>Original values remain available in review and audit exports.</small></label></div><label className="check"><input type="checkbox" checked={options.requireName} onChange={e => option('requireName', e.target.checked)} />Require a first name</label><label className="check"><input type="checkbox" checked={options.splitNames} onChange={e => option('splitNames', e.target.checked)} />Split full names into first and last name</label><p className="hint">Business names with common company suffixes stay together. Turn splitting off for other business names or names you want to preserve.</p></section>
        <section className="card"><h2>Keep additional columns</h2><p className="hint">Selected columns retain their complete original text, including commas and line breaks.</p><div className="extra-grid">{data.headers.map((h, i) => <label className="check" key={i}><input type="checkbox" checked={options.extra.includes(i)} onChange={e => option('extra', e.target.checked ? [...options.extra, i] : options.extra.filter(n => n !== i))} />{i + 1}. {h}</label>)}</div></section>
        {mappingError && <p role="alert" className="error">{mappingError}</p>}<div className="action-row"><button className="btn btn-primary" disabled={busy || !!mappingError || !data.rows.length} onClick={run}>Clean and review →</button></div>
      </> : <>
        <div className="review-title"><div><h2>{stage === 3 ? 'Your export is ready' : 'Review your cleaned list'}</h2><p className="hint">Your original file is unchanged. Review flagged rows before using your list.</p></div><button className="btn btn-ghost" disabled={busy} onClick={() => setStage(1)}>Change settings</button></div>
        <div className="stats-grid">{[[results.length, 'Original contacts'], [kept.length, 'Ready to export'], [removed.length, 'Removed'], [results.filter(r => r.warnings.length).length, 'With warnings']].map(([n, label]) => <div className="stat" key={label}><span className="stat-num">{n.toLocaleString()}</span><span className="stat-label">{label}</span></div>)}</div>
        <section className="card review-card"><div className="toolbar"><label className="search">Search contacts<input type="search" placeholder="Name, email, phone, or reason…" value={query} onChange={e => { setQuery(e.target.value); setPage(0) }} /></label><label>Show<select value={filter} onChange={e => { setFilter(e.target.value); setPage(0) }}><option value="all">All contacts</option><option value="kept">Kept</option><option value="removed">Removed</option><option value="warnings">With warnings</option></select></label></div><div className="table-wrap"><table><thead><tr><th>Source row</th><th>Contact</th><th>Phone</th><th>Email</th><th>Status / reason</th><th>Actions</th></tr></thead><tbody>{filtered.slice(currentPage * PAGE_SIZE, (currentPage + 1) * PAGE_SIZE).map(r => <tr key={r.id}><td>{r.id}</td><td>{r.output.first_name} {r.output.last_name}</td><td>{r.output.phone || '—'}{r.output.phone_extension && ` ext ${r.output.phone_extension}`}</td><td>{r.output.email || '—'}</td><td><span className={`badge ${r.reason && !r.restored ? 'removed' : ''}`}>{r.restored ? 'Restored' : r.reason || 'Kept'}</span>{r.warnings.length > 0 && <small className="warning">{r.warnings.join(' · ')}</small>}</td><td><button className="text-btn" onClick={() => setDetails(r.id)}>Compare</button>{r.reason && <button className="text-btn" onClick={() => setResults(prev => prev.map(item => item.id === r.id ? { ...item, restored: !item.restored } : item))}>{r.restored ? 'Undo restore' : 'Restore'}</button>}</td></tr>)}</tbody></table>{!filtered.length && <p className="empty">No contacts match this view.</p>}</div><div className="pagination"><span>{filtered.length} contacts · Page {currentPage + 1} of {pages}</span><button className="btn btn-ghost" disabled={currentPage === 0} onClick={() => setPage(currentPage - 1)}>Previous</button><button className="btn btn-ghost" disabled={currentPage >= pages - 1} onClick={() => setPage(currentPage + 1)}>Next</button></div></section>
        {details !== null && (() => { const row = results.find(r => r.id === details); return <section className="card"><div className="review-title"><h2>Source row {details}: before and after</h2><button className="btn btn-ghost" onClick={() => setDetails(null)}>Close comparison</button></div><div className="two-col"><div><h3>Original values</h3><dl>{data.headers.map((h, i) => <div key={i}><dt>{h}</dt><dd>{row.cells[i] || '—'}</dd></div>)}</dl></div><div><h3>Export values</h3><dl>{Object.entries(row.output).map(([k, v]) => <div key={k}><dt>{k}</dt><dd>{v || '—'}</dd></div>)}</dl></div></div>{row.reason && <p className="notice">{row.reason}. Restoring includes this row even if it fails a rule; invalid phone/email values stay blank. Original values are preserved in the audit export.</p>}</section> })()}
        <section className="card"><h2>Export your results</h2><p className="hint">{kept.length} contacts ready. CSV uses an apostrophe before formula-like values (including + phones) for spreadsheet safety. XLSX preserves phones as text without that prefix.</p><div className="action-row"><button className="btn btn-primary" disabled={busy || !kept.length} onClick={() => download('cleaned', 'xlsx')}>Export XLSX</button><button className="btn btn-secondary" disabled={busy || !kept.length} onClick={() => download('cleaned')}>Export CSV</button><button className="btn btn-ghost" disabled={busy || !removed.length} onClick={() => download('removed')}>Removed rows + reasons</button><button className="btn btn-ghost" disabled={busy || !results.length} onClick={() => download('audit')}>Full audit CSV</button></div></section>
      </>}
      </>}
      <footer>Processed on this device · No contact uploads · Original files stay untouched</footer>
    </main>
  </div>
}
