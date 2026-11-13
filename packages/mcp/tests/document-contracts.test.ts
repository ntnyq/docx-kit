import Ajv from 'ajv'
import { expect, it } from 'vitest'
import { BUILTIN_PLUGIN_CATALOG } from '../src/plugins/catalog'
import { docxSchemaResource } from '../src/resources/schema'
import { validateSchema } from '../src/tools/validateSchema'

const validateJson = new Ajv({ allErrors: true, strict: false }).compile(
  docxSchemaResource.schema,
)

it.each(
  BUILTIN_PLUGIN_CATALOG.filter(
    plugin =>
      !['bibliography', 'divider', 'pageNumber', 'toc'].includes(plugin.name),
  ),
)(
  'rejects missing required content for $name in both validators',
  ({ name }) => {
    const schema = { content: [{ name, options: {}, type: 'plugin' }] }
    expect(validateSchema(schema).valid).toBe(false)
    expect(validateJson(schema)).toBe(false)
  },
)

it.each([
  { page: { orientation: 'sideways' } },
  { page: { size: { height: 20, width: 'oops' } } },
  { page: { size: { width: '20mm' } } },
  { page: { margin: {} } },
  { plugins: {} },
  { plugins: [{ type: 'npm' }] },
  { plugins: [{ path: '', type: 'local' }] },
  { plugins: [{ type: 'url', url: 1 }] },
  { plugins: [{ plugin: {}, type: 'inline' }] },
  { plugins: [{ type: 'unknown' }] },
])('rejects invalid document options %j consistently', options => {
  const schema = { content: [], ...options }
  expect(validateSchema(schema).valid).toBe(false)
  expect(validateJson(schema)).toBe(false)
})

it.each([
  [
    'invoice',
    {
      date: '2026-11-13',
      from: { name: 'Seller' },
      invoiceNumber: '1',
      items: [{ description: 'Item', quantity: 0, unitPrice: 1 }],
      to: { name: 'Buyer' },
    },
  ],
  ['checklist', { items: [{}] }],
  ['meetingMinutes', { agenda: [{}], attendees: [1], date: '', title: '' }],
  ['metricCard', { items: [{ label: 'Revenue', value: {} }] }],
])('rejects malformed nested %s content', (name, options) => {
  const schema = { content: [{ name, options, type: 'plugin' }] }
  expect(validateSchema(schema).valid).toBe(false)
  expect(validateJson(schema)).toBe(false)
})

it('keeps custom plugin options extensible and accepts supported page lengths', () => {
  const schema = {
    content: [{ name: 'custom', options: { arbitrary: true }, type: 'plugin' }],
    plugins: [{ package: 'custom', type: 'npm' }],
    page: {
      margin: '10mm 20mm',
      orientation: 'landscape',
      size: { height: '200mm', width: '300mm' },
    },
  }
  expect(validateSchema(schema).valid).toBe(true)
  expect(validateJson(schema)).toBe(true)
})
