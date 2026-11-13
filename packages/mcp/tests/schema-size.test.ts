import { Buffer } from 'node:buffer'
import { generateToolDefinitions } from '@docxkit/ai'
import Ajv from 'ajv'
import { expect, it } from 'vitest'
import { docxSchemaResource } from '../src/resources/schema'

// Bound schema expansion rather than wall-clock time: repeated style fragments
// previously expanded these roots beyond 500 KiB and made cold compilation slow.
it.each([
  ['MCP resource', docxSchemaResource.schema],
  [
    'AI tool',
    generateToolDefinitions().find(
      tool => tool.function.name === 'create_document',
    )!.function.parameters,
  ],
])(
  'keeps the %s schema compact without dropping plugin contracts',
  (_name, schema) => {
    const serialized = JSON.stringify(schema)
    expect(Buffer.byteLength(serialized)).toBeLessThan(100 * 1024)
    const validate = new Ajv({ strict: false }).compile(JSON.parse(serialized))
    expect(
      validate({
        content: [
          {
            name: 'qrcode',
            type: 'plugin',
            options: {
              styles: { caption: { fontSize: '12pt' } },
              text: 'Hello',
            },
          },
        ],
      }),
    ).toBe(true)
    expect(
      validate({
        content: [
          {
            name: 'qrcode',
            type: 'plugin',
            options: {
              styles: { caption: { fontSize: true } },
              text: 'Hello',
            },
          },
        ],
      }),
    ).toBe(false)
  },
)
