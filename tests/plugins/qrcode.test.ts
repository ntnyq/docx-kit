import { describe, expect, it, vi } from 'vitest'

import { qrcodePlugin } from '../../src/plugins/qrcode'

// Mock qrcode
vi.mock('qrcode', () => ({
  toDataURL: vi.fn().mockResolvedValue('data:image/png;base64,iVBORw0KGgo='),
  default: {
    toDataURL: vi.fn().mockResolvedValue('data:image/png;base64,iVBORw0KGgo='),
  },
}))

describe('qrcodePlugin', () => {
  it('returns a plugin with name "qrcode"', () => {
    const plugin = qrcodePlugin()
    expect(plugin.name).toBe('qrcode')
  })

  it('renders a QR code with default options', async () => {
    const plugin = qrcodePlugin()
    const result = await plugin.render({ text: 'https://example.com' }, {
      config: {},
      utils: {
        image: {
          fromBlob: async () => new Uint8Array(),
          fromDataUrl: async () => new Uint8Array([1, 2, 3]),
        },
      },
      compileNode: async () => null,
    } as any)
    expect(result).toBeDefined()
    expect(Array.isArray(result)).toBe(true)
  })

  it('renders a QR code with caption', async () => {
    const plugin = qrcodePlugin()
    const result = await plugin.render(
      { caption: 'Scan me', text: 'https://example.com' },
      {
        config: {},
        utils: {
          image: {
            fromBlob: async () => new Uint8Array(),
            fromDataUrl: async () => new Uint8Array([1, 2, 3]),
          },
        },
        compileNode: async () => null,
      } as any,
    )
    expect(Array.isArray(result)).toBe(true)
    // Two paragraphs: image + caption
    expect((result as any[]).length).toBe(2)
  })

  it('renders with custom size and error correction', async () => {
    const plugin = qrcodePlugin()
    const result = await plugin.render(
      {
        errorCorrectionLevel: 'H',
        margin: 4,
        size: 256,
        text: 'https://example.com',
      },
      {
        config: {},
        utils: {
          image: {
            fromBlob: async () => new Uint8Array(),
            fromDataUrl: async () => new Uint8Array([1, 2, 3]),
          },
        },
        compileNode: async () => null,
      } as any,
    )
    expect(result).toBeDefined()
  })
})
