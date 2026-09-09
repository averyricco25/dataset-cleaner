import test from 'node:test'
import assert from 'node:assert/strict'
import { cleanRows, sheetData, suggestMapping } from '../src/cleaner.js'
const headers = ['first_name', 'last_name', 'phone', 'email', 'notes']
const defaults = { contact: 'either', country: 'US', requireName: true, splitNames: true, multi: 'valid', dedupe: 'either', keep: 'complete', extra: [4] }
const run = (cells, options = {}) => cleanRows(cells.map((cells, i) => ({ id: i + 2, cells })), headers, suggestMapping(headers), { ...defaults, ...options })
test('preserves notes and extension while rejecting invalid contact values', () => {
  const [row, invalid] = run([['Alex', 'West', '3125550101 ext 123', 'a@example.com', 'VIP, repeat; customer\nline 2'], ['Sam', 'Lee', '0000000000', 'not-an-email', '']])
  assert.equal(row.output.phone, '+13125550101'); assert.equal(row.output.phone_extension, '123')
  assert.equal(row.output.notes, 'VIP, repeat; customer\nline 2'); assert.equal(row.reason, '')
  assert.equal(invalid.reason, 'No valid phone or email'); assert.equal(invalid.warnings.length, 2)
})
test('deduplicates email case-insensitively and retains most complete row', () => {
  const result = run([['Alex', '', '', 'A@example.com', ''], ['Alex', 'West', '', 'a@example.com', 'details']])
  assert.equal(result[0].reason, 'Duplicate of source row 3'); assert.equal(result[1].reason, '')
  assert.equal(run([['Alex', '', '', 'a@example.com', ''], ['Alex', '', '', 'a@example.com', '']], { dedupe: 'none' }).filter(r => !r.reason).length, 2)
})
test('separates extensions and supports first-valid selection', () => {
  const result = run([['Alex', 'West', 'bad;3125550101 ext 1', '', ''], ['Sam', 'West', '3125550101 ext 2', '', '']])
  assert.ok(result.every(r => !r.reason)); assert.equal(result[0].output.phone_extension, '1')
  assert.equal(run([['Alex', '', 'bad;3125550101', '', '']], { multi: 'first' })[0].reason, 'No valid phone or email')
})
test('never interprets first contact as a header, preserves duplicate column positions', () => {
  const data = sheetData([['name', 'name', 'phone'], ['Mobile Services', 'Other', '3125550101']])
  assert.equal(data.rows.length, 1); assert.equal(data.rows[0].id, 2); assert.equal(data.rows[0].cells[0], 'Mobile Services')
  assert.equal(suggestMapping(['last_contact', 'phone_notes', 'email']) .phone, '-1')
})
test('supports business name preservation and optional name requirements', () => {
  const mapping = { first_name: '-1', last_name: '-1', full_name: '0', phone: '1', email: '-1' }
  const result = cleanRows([{ id: 2, cells: ['River House', '3125550101'] }], ['name', 'phone'], mapping, { ...defaults, splitNames: false, extra: [] })
  assert.equal(result[0].output.first_name, 'River House'); assert.equal(result[0].output.last_name, '')
  assert.equal(run([['', '', '3125550101', '', '']], { requireName: false })[0].reason, '')
})
