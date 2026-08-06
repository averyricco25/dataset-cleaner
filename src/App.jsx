import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const CORE_COLS = ['first_name', 'last_name', 'phone', 'email']

function normalizeHeader(h) {
  return h.toLowerCase().trim().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '')
}

// Find the first header key that matches any of the given patterns
function findCol(headers, patterns) {
  return headers.find(h => patterns.some(p => h === p || h.includes(p)))
}

function formatE164(raw) {
  if (!raw) return null
  const digits = String(raw).replace(/\D/g, '')
  // Only accept valid US/Canada phone numbers
  if (digits.length === 10) return '+1' + digits
  if (digits.length === 11 && digits[0] === '1') return '+' + digits
  return null
}

function firstValue(val) {
  if (val == null) return ''
  const s = String(val).trim()
  return s.split(/[;|,]/)[0].trim()
}

function normalizeNameCase(name) {
  if (!name) return name
  const letters = name.replace(/[^a-z]/gi, '')
  if (!letters) return name
  const isAllCaps = letters === letters.toUpperCase()
  if (!isAllCaps) return name
  return name.split(/\s+/).map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ')
}

function parseName(row, normHeaders, rawHeaders) {
  const fnKey = findCol(normHeaders, ['first_name', 'firstname', 'first', 'fname', 'given_name'])
  const lnKey = findCol(normHeaders, ['last_name', 'lastname', 'last', 'lname', 'surname', 'family_name'])

  // If both first and last name columns exist AND have values, use them
  if (fnKey && lnKey && firstValue(row[fnKey])) {
    return { first_name: normalizeNameCase(firstValue(row[fnKey])), last_name: normalizeNameCase(firstValue(row[lnKey])) }
  }

  // If only first name exists and has value, use it with empty last name
  if (fnKey && firstValue(row[fnKey]) && !lnKey) {
    return { first_name: normalizeNameCase(firstValue(row[fnKey])), last_name: '' }
  }

  // Try to find a display/business/organization name column as fallback
  // More specific patterns first, and exclude fnKey/lnKey to avoid matching those again
  const fullKey = findCol(normHeaders.filter((h, i) => h !== fnKey && h !== lnKey), ['display_name', 'displayname', 'organization_name', 'business_name', 'org_name', 'company_name', 'business', 'full_name', 'fullname', 'contact_name', 'customer_name', 'client_name', 'name'])
  if (fullKey && row[fullKey]) {
    const full = String(row[fullKey]).trim()

    // Check if it's a comma-separated name (likely "Last, First")
    if (full.includes(',')) {
      const [last, first] = full.split(',').map(s => s.trim())
      return { first_name: normalizeNameCase(first || ''), last_name: normalizeNameCase(last || '') }
    }

    // Check if it looks like a business name (has business keywords)
    const businessKeywords = /\b(inc|llc|corp|co\.|ltd|company|contractors|services|group|associates|partners|llp|pllc|agency|solutions|systems|networks|studios|works|house|store|shop|market|center|park|plaza|village)\b/i
    const hasBusinessKeyword = businessKeywords.test(full)
    const words = full.trim().split(/\s+/)

    // If it has business keywords, treat as business name
    if (hasBusinessKeyword) {
      return { first_name: normalizeNameCase(full), last_name: '' }
    }

    // For names with multiple words, split into first and last
    // The last word becomes last_name, everything else (including "&") is first_name
    if (words.length >= 2) {
      const lastName = words[words.length - 1]
      const firstName = words.slice(0, -1).join(' ')
      return { first_name: normalizeNameCase(firstName), last_name: normalizeNameCase(lastName) }
    }

    // Single word - default to first name only
    return { first_name: normalizeNameCase(full), last_name: '' }
  }

  return { first_name: '', last_name: '' }
}

function getExtraCols(rawHeaders, normHeaders) {
  const fnKey = findCol(normHeaders, ['first_name', 'firstname', 'first', 'fname', 'given_name'])
  const lnKey = findCol(normHeaders, ['last_name', 'lastname', 'last', 'lname', 'surname', 'family_name'])
  const fullKey = findCol(
    normHeaders.filter(h => h !== fnKey && h !== lnKey),
    ['display_name', 'displayname', 'organization_name', 'business_name', 'org_name', 'company_name', 'business', 'full_name', 'fullname', 'contact_name', 'customer_name', 'client_name', 'name']
  )
  const phoneKey = findCol(normHeaders, ['phone', 'phone_number', 'phonenumber', 'mobile', 'cell', 'telephone', 'cell_phone', 'mobile_phone', 'contact_phone'])
  const emailKey = findCol(normHeaders, ['email', 'email_address', 'emailaddress', 'e_mail'])
  const coreNormKeys = new Set([fnKey, lnKey, fullKey, phoneKey, emailKey].filter(Boolean))
  return rawHeaders
    .map((raw, i) => ({ raw, norm: normHeaders[i] }))
    .filter(({ norm }) => !coreNormKeys.has(norm))
}

function cleanData(rows, rawHeaders, extraColsToKeep, prioritizeEmail) {
  const normHeaders = rawHeaders.map(normalizeHeader)

  const summary = {
    started: rows.length,
    missingName: 0,
    missingPhone: 0,
    badPhone: 0,
    duplicate: 0,
    emailOnly: 0,
    final: 0,
  }

  const seenPhones = new Set()
  const cleaned = []
  const removed = []

  // Find phone and email columns once
  const phoneKey = findCol(normHeaders, ['phone', 'phone_number', 'phonenumber', 'mobile', 'cell', 'telephone', 'cell_phone', 'mobile_phone', 'contact_phone'])
  const emailKey = findCol(normHeaders, ['email', 'email_address', 'emailaddress', 'e_mail'])

  for (const rawRow of rows) {
    const row = {}
    rawHeaders.forEach((h, i) => { row[normHeaders[i]] = rawRow[h] })

    const { first_name, last_name } = parseName(row, normHeaders, rawHeaders)

    if (!first_name) {
      summary.missingName++
      removed.push({ ...row, _reason: 'Missing first name' })
      continue
    }

    const phoneRaw = phoneKey ? firstValue(row[phoneKey]) : ''
    const phone = formatE164(phoneRaw)
    const emailRaw = emailKey ? (row[emailKey] || '') : ''
    const email = firstValue(emailRaw)

    const hasPhone = !!phoneRaw
    const hasValidPhone = !!phone
    const hasEmail = !!email

    // Check phone requirement
    if (!hasPhone) {
      if (prioritizeEmail && hasEmail) {
        // Email-only contact is allowed
        summary.emailOnly++
      } else {
        summary.missingPhone++
        removed.push({ ...row, _reason: 'Missing phone and email' })
        continue
      }
    } else if (!hasValidPhone) {
      if (prioritizeEmail && hasEmail) {
        // Invalid phone but has email - keep it
        summary.emailOnly++
      } else {
        summary.badPhone++
        removed.push({ ...row, _reason: 'Invalid phone format' })
        continue
      }
    } else if (seenPhones.has(phone)) {
      summary.duplicate++
      removed.push({ ...row, _reason: 'Duplicate phone' })
      continue
    } else {
      seenPhones.add(phone)
    }

    const out = { first_name, last_name, phone, email }

    extraColsToKeep.forEach(col => {
      out[col.raw] = firstValue(String(row[col.norm] ?? ''))
    })

    cleaned.push(out)
  }

  summary.final = cleaned.length
  return { cleaned, removed, summary }
}

function exportCSV(data, filename) {
  const ws = XLSX.utils.json_to_sheet(data)
  const csv = XLSX.utils.sheet_to_csv(ws)
  const blob = new Blob([csv], { type: 'text/csv' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

function exportXLSX(data, filename) {
  const ws = XLSX.utils.json_to_sheet(data)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Cleaned')
  XLSX.writeFile(wb, filename)
}

export default function App() {
  const [stage, setStage] = useState('upload') // upload | extra-cols | result
  const [dragOver, setDragOver] = useState(false)
  const [fileName, setFileName] = useState('')
  const [selectedFile, setSelectedFile] = useState(null)
  const [rawRows, setRawRows] = useState([])
  const [rawHeaders, setRawHeaders] = useState([])
  const [extraCols, setExtraCols] = useState([])
  const [selectedExtraCols, setSelectedExtraCols] = useState(new Set())
  const [prioritizeEmail, setPrioritizeEmail] = useState(false)
  const [cleanedRows, setCleanedRows] = useState([])
  const [removedRows, setRemovedRows] = useState([])
  const [summary, setSummary] = useState(null)
  const [showRemoved, setShowRemoved] = useState(false)

  const parseFile = useCallback((file) => {
    setFileName(file.name)
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result)
      const wb = XLSX.read(data, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      let rows = XLSX.utils.sheet_to_json(ws, { defval: '' })
      if (!rows.length) return

      let headers = Object.keys(rows[0])

      // Check if first data row looks like column headers
      const firstRow = rows[0]
      const firstRowValues = Object.values(firstRow)
      const looksLikeHeaders = firstRowValues.some(v => /^(first|last|name|phone|email|mobile|contact)/i.test(String(v).trim()))

      if (looksLikeHeaders) {
        // Use first row values as headers and remove it from data
        headers = firstRowValues.map(v => String(v).trim())
        rows = rows.slice(1).map(row => {
          const newRow = {}
          headers.forEach((h, i) => {
            const oldKey = Object.keys(row)[i]
            newRow[h] = row[oldKey] || ''
          })
          return newRow
        })
      }

      // Store file data but don't process yet
      setSelectedFile({ rows, headers })
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const processFile = () => {
    if (!selectedFile || !fileName) {
      return
    }

    const { rows, headers } = selectedFile
    setRawRows(rows)
    setRawHeaders(headers)

    const normHeaders = headers.map(normalizeHeader)
    const cols = getExtraCols(headers, normHeaders)
    setExtraCols(cols)

    if (cols.length > 0) {
      setSelectedExtraCols(new Set(cols.map(c => c.norm)))
      setStage('extra-cols')
    } else {
      const { cleaned, removed, summary } = cleanData(rows, headers, [], prioritizeEmail)
      setCleanedRows(cleaned)
      setRemovedRows(removed)
      setSummary(summary)
      setStage('result')
    }
  }

  const onDrop = useCallback((e) => {
    e.preventDefault()
    setDragOver(false)
    const file = e.dataTransfer.files[0]
    if (file) parseFile(file)
  }, [parseFile])

  const onFileInput = (e) => {
    const file = e.target.files[0]
    if (file) parseFile(file)
  }

  const confirmExtraCols = (overrideSelection) => {
    const sel = overrideSelection ?? selectedExtraCols
    const toKeep = extraCols.filter(c => sel.has(c.norm))
    const { cleaned, removed, summary } = cleanData(rawRows, rawHeaders, toKeep, prioritizeEmail)
    setCleanedRows(cleaned)
    setRemovedRows(removed)
    setSummary(summary)
    setStage('result')
  }

  const handleExport = () => {
    const base = fileName.replace(/\.[^.]+$/, '')
    exportCSV(cleanedRows, `${base}_cleaned.csv`)
  }

  const handleExportXLSX = () => {
    const base = fileName.replace(/\.[^.]+$/, '')
    exportXLSX(cleanedRows, `${base}_cleaned.xlsx`)
  }

  const reset = () => {
    setStage('upload')
    setFileName('')
    setRawRows([])
    setRawHeaders([])
    setExtraCols([])
    setSelectedExtraCols(new Set())
    setCleanedRows([])
    setRemovedRows([])
    setSummary(null)
    setShowRemoved(false)
  }

  return (
    <div className="app">
      <header className="header">
        <div className="header-inner">
          <span className="logo-mark">◈</span>
          <h1>Contact List Cleaner</h1>
          <span className="subtitle">for Relentless Digital</span>
        </div>
      </header>

      <main className="main">
        {stage === 'upload' && (
          <div className="card upload-card">
            <h2>Upload your contact list</h2>
            <p className="hint">Accepts CSV, XLSX, or XLS</p>
            <div
              className={`dropzone${dragOver ? ' dragover' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => document.getElementById('file-input').click()}
            >
              {selectedFile && fileName ? (
                <>
                  <div className="drop-icon" style={{ fontSize: '2rem', color: '#4caf50' }}>✓</div>
                  <p style={{ margin: '0.5rem 0 0 0', fontWeight: 500 }}>{fileName}</p>
                  <p className="drop-sub">Ready to clean</p>
                </>
              ) : (
                <>
                  <div className="drop-icon">↑</div>
                  <p>Drag & drop your file here</p>
                  <p className="drop-sub">or click to browse</p>
                </>
              )}
              <input
                id="file-input"
                type="file"
                accept=".csv,.xlsx,.xls"
                style={{ display: 'none' }}
                onChange={onFileInput}
              />
            </div>
            <div className="rules-list">
              <h3>Cleaning rules applied</h3>
              <ul>
                <li>Split full names into first / last columns</li>
                <li>Format phones to E.164 (+1XXXXXXXXXX)</li>
                <li>Keep only first value in multi-value cells</li>
                <li>Remove rows missing first name or phone</li>
                <li>Deduplicate by phone (keep first occurrence)</li>
                <li>Drop irrelevant columns</li>
              </ul>
            </div>

            <div style={{ marginTop: '2rem', paddingTop: '2rem', borderTop: '1px solid rgba(255,255,255,0.1)' }}>
              <p className="tag-question">Contact method priority:</p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: '0.75rem', cursor: 'pointer', padding: '0.75rem', borderRadius: '6px', backgroundColor: !prioritizeEmail ? 'rgba(76, 175, 80, 0.1)' : 'transparent', border: !prioritizeEmail ? '1px solid rgba(76, 175, 80, 0.3)' : '1px solid rgba(255,255,255,0.1)', transition: 'all 0.2s' }}>
                  <input
                    type="radio"
                    name="contactMethod"
                    checked={!prioritizeEmail}
                    onChange={() => setPrioritizeEmail(false)}
                    style={{ marginTop: '0.25rem', cursor: 'pointer' }}
                  />
                  <div>
                    <p style={{ margin: '0 0 0.25rem 0', fontWeight: 500 }}>Phone required</p>
                    <p className="hint" style={{ margin: 0 }}>Removes any contacts missing a valid phone number</p>
                  </div>
                </label>
                <label style={{ display: 'flex', alignItems: 'flex-start', gap: '0.75rem', cursor: 'pointer', padding: '0.75rem', borderRadius: '6px', backgroundColor: prioritizeEmail ? 'rgba(76, 175, 80, 0.1)' : 'transparent', border: prioritizeEmail ? '1px solid rgba(76, 175, 80, 0.3)' : '1px solid rgba(255,255,255,0.1)', transition: 'all 0.2s' }}>
                  <input
                    type="radio"
                    name="contactMethod"
                    checked={prioritizeEmail}
                    onChange={() => setPrioritizeEmail(true)}
                    style={{ marginTop: '0.25rem', cursor: 'pointer' }}
                  />
                  <div>
                    <p style={{ margin: '0 0 0.25rem 0', fontWeight: 500 }}>Prioritize email</p>
                    <p className="hint" style={{ margin: 0 }}>Keeps contacts that have email addresses, even without phone numbers (useful for email campaigns)</p>
                  </div>
                </label>
              </div>
            </div>

            <div style={{ marginTop: '2rem', paddingTop: '2rem', borderTop: '1px solid rgba(255,255,255,0.1)' }}>
              <button
                className={`btn btn-lg ${selectedFile ? 'btn-primary' : 'btn-ghost'}`}
                onClick={processFile}
                disabled={!selectedFile}
                style={{
                  width: '100%',
                  opacity: selectedFile ? 1 : 0.4,
                  cursor: selectedFile ? 'pointer' : 'not-allowed',
                  pointerEvents: selectedFile ? 'auto' : 'none'
                }}
              >
                Clean this file
              </button>
            </div>
          </div>
        )}

        {stage === 'extra-cols' && (
          <div className="card tag-card">
            <h2>Additional columns detected</h2>
            <p className="hint">These columns aren't part of the standard output. Select the ones you'd like to keep:</p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem', margin: '1.5rem 0' }}>
              {extraCols.map(col => (
                <label key={col.norm} style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', cursor: 'pointer', padding: '0.6rem 0.75rem', borderRadius: '6px', backgroundColor: selectedExtraCols.has(col.norm) ? 'rgba(76, 175, 80, 0.1)' : 'transparent', border: selectedExtraCols.has(col.norm) ? '1px solid rgba(76, 175, 80, 0.3)' : '1px solid rgba(255,255,255,0.1)', transition: 'all 0.2s' }}>
                  <input
                    type="checkbox"
                    checked={selectedExtraCols.has(col.norm)}
                    onChange={(e) => {
                      setSelectedExtraCols(prev => {
                        const next = new Set(prev)
                        if (e.target.checked) next.add(col.norm)
                        else next.delete(col.norm)
                        return next
                      })
                    }}
                    style={{ cursor: 'pointer', width: '1rem', height: '1rem' }}
                  />
                  <span style={{ fontWeight: 500 }}>{col.raw}</span>
                </label>
              ))}
            </div>
            <div className="btn-row">
              <button className="btn btn-primary" onClick={() => confirmExtraCols()}>Continue</button>
              <button className="btn btn-ghost" onClick={() => confirmExtraCols(new Set())}>Keep none</button>
            </div>
          </div>
        )}

        {stage === 'result' && summary && (
          <div className="results">
            <div className="card summary-card">
              <h2>Cleaning Summary</h2>
              <div className="stats-grid">
                <div className="stat">
                  <span className="stat-num">{summary.started}</span>
                  <span className="stat-label">Rows started</span>
                </div>
                <div className="stat stat-removed">
                  <span className="stat-num">{summary.started - summary.final}</span>
                  <span className="stat-label">Rows removed</span>
                </div>
                <div className="stat stat-final">
                  <span className="stat-num">{summary.final}</span>
                  <span className="stat-label">Final rows</span>
                </div>
              </div>
              <div className="breakdown">
                <h3>Removed breakdown</h3>
                <div className="breakdown-rows">
                  {summary.missingName > 0 && (
                    <div className="breakdown-row">
                      <span>Missing first name</span>
                      <span className="breakdown-num">{summary.missingName}</span>
                    </div>
                  )}
                  {summary.missingPhone > 0 && (
                    <div className="breakdown-row">
                      <span>Missing phone and email</span>
                      <span className="breakdown-num">{summary.missingPhone}</span>
                    </div>
                  )}
                  {summary.badPhone > 0 && (
                    <div className="breakdown-row">
                      <span>Invalid phone format</span>
                      <span className="breakdown-num">{summary.badPhone}</span>
                    </div>
                  )}
                  {summary.duplicate > 0 && (
                    <div className="breakdown-row">
                      <span>Duplicate phone</span>
                      <span className="breakdown-num">{summary.duplicate}</span>
                    </div>
                  )}
                  {(summary.missingName + summary.missingPhone + summary.badPhone + summary.duplicate) === 0 && (
                    <div className="breakdown-row">
                      <span>No rows removed</span>
                      <span className="breakdown-num">0</span>
                    </div>
                  )}
                  {(summary.missingName + summary.missingPhone + summary.badPhone + summary.duplicate) > 0 && (
                    <div className="breakdown-row" style={{ borderLeft: '3px solid #ef5350', paddingLeft: '0.75rem' }}>
                      <span>Total removed</span>
                      <span className="breakdown-num">{summary.missingName + summary.missingPhone + summary.badPhone + summary.duplicate}</span>
                    </div>
                  )}
                </div>
              </div>

              {summary.emailOnly > 0 && (
                <div className="breakdown" style={{ marginTop: '1.5rem' }}>
                  <h3>Kept (email-only)</h3>
                  <p className="hint" style={{ marginBottom: '1rem' }}>These contacts were kept because they have email addresses but no phone numbers. This is due to the "Prioritize email" option you selected.</p>
                  <div className="breakdown-rows">
                    <div className="breakdown-row" style={{ color: '#64b5f6' }}>
                      <span>Email-only contacts (no phone)</span>
                      <span className="breakdown-num">{summary.emailOnly}</span>
                    </div>
                  </div>
                </div>
              )}
            </div>

            <div className="card preview-card">
              <h2>Preview <span className="preview-note">(first 10 rows)</span></h2>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {cleanedRows.length > 0 && Object.keys(cleanedRows[0]).map(k => (
                        <th key={k}>{k}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {cleanedRows.slice(0, 10).map((row, i) => (
                      <tr key={i}>
                        {Object.values(row).map((v, j) => (
                          <td key={j}>{v}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {removedRows.length > 0 && (
              <div className="card removed-card">
                <div className="removed-header" onClick={() => setShowRemoved(v => !v)}>
                  <h2>Removed Rows <span className="preview-note">({removedRows.length} total)</span></h2>
                  <span className="toggle-icon">{showRemoved ? '▲' : '▼'}</span>
                </div>
                {showRemoved && (
                  <div className="table-wrap">
                    <table>
                      <thead>
                        <tr>
                          <th>Reason</th>
                          {Object.keys(removedRows[0]).filter(k => k !== '_reason').slice(0, 5).map(k => (
                            <th key={k}>{k}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {removedRows.slice(0, 10).map((row, i) => (
                          <tr key={i} className="removed-row">
                            <td><span className="reason-badge">{row._reason}</span></td>
                            {Object.entries(row).filter(([k]) => k !== '_reason').slice(0, 5).map(([k, v]) => (
                              <td key={k}>{String(v)}</td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {removedRows.length > 10 && (
                      <p className="table-more">…and {removedRows.length - 10} more removed rows</p>
                    )}
                  </div>
                )}
              </div>
            )}

            <div className="action-row">
              <button className="btn btn-primary btn-lg" onClick={handleExport}>
                ↓ Export CSV
              </button>
              <button className="btn btn-secondary btn-lg" onClick={handleExportXLSX}>
                ↓ Export XLSX
              </button>
              <button className="btn btn-ghost" onClick={reset}>
                Clean another file
              </button>
            </div>
          </div>
        )}
      </main>
    </div>
  )
}
