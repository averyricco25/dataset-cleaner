import { parsePhoneNumberFromString } from 'libphonenumber-js/max'

export const FIELDS = { first_name: 'First name', last_name: 'Last name', full_name: 'Full or business name', phone: 'Phone', email: 'Email' }
const aliases = { first_name: ['first_name', 'firstname', 'first', 'fname', 'given_name'], last_name: ['last_name', 'lastname', 'last', 'surname', 'family_name'], full_name: ['full_name', 'fullname', 'name', 'display_name', 'business_name', 'company_name', 'contact_name'], phone: ['phone', 'phone_number', 'mobile', 'cell', 'telephone', 'mobile_phone'], email: ['email', 'email_address', 'e_mail'] }
export const text = value => String(value ?? '').trim()
export function suggestMapping(headers) {
  const normalized = headers.map(h => text(h).toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, ''))
  return Object.fromEntries(Object.entries(aliases).map(([key, names]) => [key, String(normalized.findIndex(h => names.includes(h)))]))
}
export function sheetData(matrix, headerRow = 0) {
  const width = matrix.reduce((width, row) => Math.max(width, row.length), 0)
  const headers = Array.from({ length: width }, (_, i) => text(matrix[headerRow]?.[i]) || `Column ${i + 1}`)
  return { headers, rows: matrix.slice(headerRow + 1).map((cells, i) => ({ id: i + headerRow + 2, cells: Array.from({ length: width }, (_, j) => String(cells[j] ?? '')) })).filter(row => row.cells.some(v => text(v))) }
}
function nameParts(full, split) {
  if (!split || /\b(inc|llc|corp|ltd|company|services|group|partners|agency|solutions)\b/i.test(full)) return [full, '']
  if (full.includes(',')) { const [last, ...first] = full.split(','); return [first.join(',').trim(), last.trim()] }
  const words = full.split(/\s+/)
  return words.length > 1 ? [words.slice(0, -1).join(' '), words.at(-1)] : [full, '']
}
export function cleanRows(rows, headers, mapping, options) {
  const get = (row, key) => text(row.cells[Number(mapping[key])])
  const results = rows.map(row => {
    const warnings = []
    let first = get(row, 'first_name'), last = get(row, 'last_name')
    if (!first && get(row, 'full_name')) [first, last] = nameParts(get(row, 'full_name'), options.splitNames)
    const rawPhone = get(row, 'phone'), rawEmail = get(row, 'email')
    const phones = rawPhone.split(/[|,;](?!ext=)/).map(text).filter(Boolean)
    const emails = rawEmail.split(/[|,;]/).map(text).filter(Boolean)
    const validPhone = value => { try { const p = parsePhoneNumberFromString(value, { defaultCountry: options.country, extract: false }); return p?.isValid() ? p : null } catch { return null } }
    const phone = (options.multi === 'first' ? phones.slice(0, 1) : phones).map(validPhone).find(Boolean)
    const email = (options.multi === 'first' ? emails.slice(0, 1) : emails).find(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) || ''
    if (rawPhone && !phone) warnings.push('Invalid phone')
    if (rawEmail && !email) warnings.push('Invalid email')
    if (phones.length > 1 || emails.length > 1) warnings.push('Multiple values: original retained in review')
    const output = { first_name: first, last_name: last, phone: phone?.number || '', phone_extension: phone?.ext || '', email }
    for (const index of options.extra) {
      let key = headers[index]
      if (Object.hasOwn(output, key)) key = `original_${key}`
      while (Object.hasOwn(output, key)) key += '_'
      output[key] = row.cells[index]
    }
    let reason = ''
    if (options.requireName && !first) reason = 'Missing first name'
    else if (options.contact === 'phone' && !phone) reason = rawPhone ? 'Invalid phone' : 'Missing phone'
    else if (options.contact === 'either' && !phone && !email) reason = 'No valid phone or email'
    else if (options.contact === 'email' && !email) reason = rawEmail ? 'Invalid email' : 'Missing email'
    return { ...row, output, warnings, reason, restored: false }
  })
  const candidates = results.filter(r => !r.reason)
  if (options.keep === 'complete') candidates.sort((a, b) => Object.values(b.output).filter(text).length - Object.values(a.output).filter(text).length || a.id - b.id)
  const seen = new Map()
  for (const row of candidates) {
    const keys = []
    if (['phone', 'either'].includes(options.dedupe) && row.output.phone) keys.push(`p:${row.output.phone}:${row.output.phone_extension}`)
    if (['email', 'either'].includes(options.dedupe) && row.output.email) keys.push(`e:${row.output.email.toLowerCase()}`)
    const duplicate = keys.map(k => seen.get(k)).find(Boolean)
    if (duplicate) row.reason = `Duplicate of source row ${duplicate}`
    else keys.forEach(k => seen.set(k, row.id))
  }
  return results
}

