import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  createDocxKitError,
  DIAGNOSTIC_CODES,
  DocxKitError,
  ERROR_CODES,
  serializeDocxKitError,
} from '../src/index'

afterEach(() => vi.restoreAllMocks())

describe('compatible structured diagnostics', () => {
  it('keeps the legacy constructor, error identity, message, code and cause', () => {
    const cause = { original: true }
    const error = new DocxKitError(
      'PLUGIN_NOT_REGISTERED',
      'Original message',
      cause,
    )
    expect(error).toBeInstanceOf(Error)
    expect(error).toBeInstanceOf(DocxKitError)
    expect(error.name).toBe('DocxKitError')
    expect(error.code).toBe('PLUGIN_NOT_REGISTERED')
    expect(error.message).toBe('Original message')
    expect(error.cause).toBe(cause)
    expect(error.diagnosticCode).toBe('DOCXKIT_PLUGIN_NOT_REGISTERED')
    expect(error.docs).toBe(
      'https://docx-kit.ntnyq.dev/errors/docxkit_plugin_not_registered/',
    )
    expect(error.fix).toContain('.use(')
  })

  it.each(['CUSTOM', '__proto__', 'constructor', 'DOCXKIT_CUSTOM'])(
    'does not invent docs for %s',
    code => {
      const error = new DocxKitError(code, 'Custom failure')
      expect(error.code).toBe(code)
      expect(error.docs).toBeUndefined()
      expect(error.fix).toBeUndefined()
      expect(error.diagnosticCode).toBeUndefined()
    },
  )

  it('creates typed diagnostics silently, without exposing Nostics error identity', () => {
    const log = vi.spyOn(console, 'log')
    const warn = vi.spyOn(console, 'warn')
    const report = vi.spyOn(console, 'error')
    const error = createDocxKitError('PLUGIN_NOT_REGISTERED', {
      pluginName: 'qrcode',
    })
    expect(error).toBeInstanceOf(DocxKitError)
    expect(error.name).toBe('DocxKitError')
    expect(error.message).toBe('Plugin not registered: qrcode')
    expect(error.diagnosticCode).toBe(DIAGNOSTIC_CODES.PLUGIN_NOT_REGISTERED)
    expect(error.stack).not.toContain('at createDocxKitError')
    expect(log).not.toHaveBeenCalled()
    expect(warn).not.toHaveBeenCalled()
    expect(report).not.toHaveBeenCalled()
  })

  it.each([undefined, null, 'raw cause', 0, false, { original: true }])(
    'preserves arbitrary causes: %s',
    cause => {
      const error = createDocxKitError(
        'EXPORT_FAILED',
        { format: 'buffer' },
        { cause },
      )
      expect(error.cause).toBe(cause)
      expect(error.message).toBe('Failed to pack document to buffer')
    },
  )

  it('accepts call-site sources and preserves explicit null over a param cause', () => {
    const error = createDocxKitError(
      'PLUGIN_RENDER_FAILED',
      {
        cause: 'inner',
        pluginName: 'example',
        sources: ['template.json:1:1'],
      },
      { cause: null, message: 'Original message' },
    )
    expect(error.cause).toBeNull()
    expect(error.sources).toEqual(['template.json:1:1'])
    expect(error.message).toBe('Original message')
  })

  it('works without the V8 stack capture API', () => {
    const capture = Error.captureStackTrace
    try {
      Reflect.set(Error, 'captureStackTrace', undefined)
      expect(createDocxKitError('EXPORT_FAILED', { format: 'blob' }).code).toBe(
        'EXPORT_FAILED',
      )
    } finally {
      Reflect.set(Error, 'captureStackTrace', capture)
    }
  })

  it('selects load repairs from an explicit reason', () => {
    const security = createDocxKitError('PLUGIN_LOAD_FAILED', {
      detail: 'Rejected',
      reason: 'security',
    })
    const environment = createDocxKitError('PLUGIN_LOAD_FAILED', {
      detail: 'Unavailable',
      reason: 'environment',
    })
    const invalidExport = createDocxKitError('PLUGIN_LOAD_FAILED', {
      detail: 'Invalid',
      reason: 'export',
    })
    expect(security.fix).toContain('keep untrusted code blocked')
    expect(environment.fix).toContain('platform-specific')
    expect(invalidExport.fix).toContain('Export a DocxPlugin')
  })

  it('maps all legacy values to full diagnostic codes', () => {
    expect(Object.values(ERROR_CODES)).toHaveLength(18)
    for (const code of Object.values(ERROR_CODES)) {
      expect(ERROR_CODES[code]).toBe(code)
      expect(DIAGNOSTIC_CODES[code]).toBe(`DOCXKIT_${code}`)
    }
  })

  it('serializes only allowlisted fields, even with circular causes', () => {
    const cause: { secret: string; self?: unknown } = {
      secret: 'private document content',
    }
    cause.self = cause
    const error = createDocxKitError(
      'PLUGIN_RENDER_FAILED',
      {
        pluginName: 'example',
        sources: ['private-template.json:1:1'],
      },
      { cause },
    )
    Object.assign(error, {
      data: { image: 'base64' },
      schema: { body: 'private' },
    })
    // Verify JSON transport behavior, including omission of undefined fields.
    // eslint-disable-next-line unicorn/prefer-structured-clone -- This assertion tests the wire format, not cloning.
    expect(JSON.parse(JSON.stringify(serializeDocxKitError(error)))).toEqual({
      code: 'PLUGIN_RENDER_FAILED',
      diagnosticCode: 'DOCXKIT_PLUGIN_RENDER_FAILED',
      docs: error.docs,
      fix: error.fix,
      message: 'Plugin render failed: example',
    })
  })

  it.each([new Error('Unexpected'), 'Unexpected', null])(
    'keeps unknown exceptions readable',
    error => {
      expect(serializeDocxKitError(error)).toEqual({
        message: error instanceof Error ? error.message : String(error),
      })
    },
  )
})

describe('safe diagnostic output regressions', () => {
  it('removes URL credentials, query strings and fragments from DTO fields only', () => {
    const url =
      'https://demo:EXAMPLE_PASSWORD@example.invalid/plugin.js?token=EXAMPLE_TOKEN#EXAMPLE_FRAGMENT'
    const message = `Failed to load "${url}" (HTTP 403)`
    const cause = { original: true }
    const error = new DocxKitError('PLUGIN_LOAD_FAILED', message, cause, {
      docs: url,
      fix: `Check ${url}`,
    })
    expect(serializeDocxKitError(error)).toMatchObject({
      code: 'PLUGIN_LOAD_FAILED',
      docs: 'https://example.invalid/plugin.js',
      fix: 'Check https://example.invalid/plugin.js',
      message: 'Failed to load "https://example.invalid/plugin.js" (HTTP 403)',
    })
    expect(error.message).toBe(message)
    expect(error.docs).toBe(url)
    expect(error.cause).toBe(cause)
  })

  it.each([
    '?redirect=https://callback.invalid/path&access_token=EXAMPLE_TOKEN',
    '?redirect=https://callback.invalid/path?nested=1&access_token=EXAMPLE_TOKEN',
    '?redirect=wss://callback.invalid/path&signature=EXAMPLE_TOKEN',
    '?redirect=file:///tmp/callback&signature=EXAMPLE_TOKEN',
    '#redirect=https://callback.invalid/path&access_token=EXAMPLE_TOKEN',
    '?token=<EXAMPLE_TOKEN>',
  ])(
    'removes the entire query or fragment from every DTO field: %s',
    suffix => {
      const url = `https://demo:EXAMPLE_PASSWORD@example.invalid/plugin.js${suffix}`
      const message = `Failed to load "${url}" (HTTP 403)`
      const error = new DocxKitError(url, message, undefined, {
        docs: url,
        fix: `Check ${url}`,
      })
      expect(serializeDocxKitError(error)).toMatchObject({
        code: 'https://example.invalid/plugin.js',
        docs: 'https://example.invalid/plugin.js',
        fix: 'Check https://example.invalid/plugin.js',
        message:
          'Failed to load "https://example.invalid/plugin.js" (HTTP 403)',
      })
      expect(serializeDocxKitError(new Error(message))).toEqual({
        message:
          'Failed to load "https://example.invalid/plugin.js" (HTTP 403)',
      })
      expect(error.code).toBe(url)
      expect(error.message).toBe(message)
      expect(error.docs).toBe(url)
      expect(error.fix).toBe(`Check ${url}`)
    },
  )

  it.each([
    'https://example.invalid/docs,https://demo:EXAMPLE_PASSWORD@example.invalid/path?token=EXAMPLE_TOKEN',
    '["https://example.invalid/docs","https://demo:EXAMPLE_PASSWORD@example.invalid/path?token=EXAMPLE_TOKEN"]',
    '<https://example.invalid/docs>,<https://demo:EXAMPLE_PASSWORD@example.invalid/path?token=EXAMPLE_TOKEN>',
  ])('sanitizes adjacent URLs independently: %s', message => {
    expect(serializeDocxKitError(new Error(message))).toEqual({
      message: message.replace(
        'https://demo:EXAMPLE_PASSWORD@example.invalid/path?token=EXAMPLE_TOKEN',
        'https://example.invalid/path',
      ),
    })
  })

  it.each([
    'https://EXAMPLE_USER:EXAMPLE_PASSWORD@example.invalid/path',
    'https://example.invalid/path?access_token=EXAMPLE_TOKEN&signature=EXAMPLE_SIGNATURE',
    'https://example.invalid/path?download=EXAMPLE_TOKEN',
    'https://example.invalid/path#EXAMPLE_FRAGMENT',
    'https://EXAMPLE_USER:EXAMPLE_PASSWORD@',
    "https://EXAMPLE_USER:EXAMPLE_PASS'WORD@example.invalid/path",
    "https://example.invalid/path?token=EXAMPLE_TO'KEN",
    'https://demo:prefix<EXAMPLE_PASSWORD>@example.invalid/path',
    'https://example.invalid/path?redirect=https://callback.invalid/path&token=EXAMPLE_TOKEN',
    'https://example.invalid/path#redirect=https://callback.invalid/path&token=EXAMPLE_TOKEN',
    'https://example.invalid/docs,https://demo:EXAMPLE_PASSWORD@example.invalid/path',
    '["https://example.invalid/docs","https://demo:EXAMPLE_PASSWORD@example.invalid/path"]',
  ])('redacts URL secrets from unexpected exceptions: %s', url => {
    const output = JSON.stringify(
      serializeDocxKitError(new Error(`Failed: ${url}`)),
    )
    expect(output).not.toContain('EXAMPLE_')
    expect(output).toContain('Failed:')
  })

  it('keeps public documentation URLs and ordinary messages intact', () => {
    const message =
      'See https://docx-kit.ntnyq.dev/errors/docxkit_export_failed/.'
    expect(serializeDocxKitError(new Error(message))).toEqual({ message })
    expect(serializeDocxKitError('HTTP 403: retry later')).toEqual({
      message: 'HTTP 403: retry later',
    })
  })

  it.each([
    'https://[::1]',
    'https://example.invalid/items[0]',
    '["https://[::1]"]',
    '<https://[::1]>',
  ])(
    'preserves balanced URL brackets and surrounding punctuation: %s',
    message => {
      expect(serializeDocxKitError(new Error(message))).toEqual({ message })
    },
  )

  it('redacts credentials and queries without removing IPv6 host brackets', () => {
    expect(
      serializeDocxKitError(
        new Error(
          'Failed: "https://demo:EXAMPLE_PASSWORD@[::1]?token=EXAMPLE_TOKEN"',
        ),
      ),
    ).toEqual({ message: 'Failed: "https://[::1]/"' })
  })

  it.each([
    Object.create(null),
    {
      toString() {
        throw new Error('Conversion failed')
      },
    },
    {
      [Symbol.toPrimitive]() {
        throw new Error('Conversion failed')
      },
    },
  ])('uses a stable fallback when an exception cannot be converted', error => {
    expect(serializeDocxKitError(error)).toEqual({ message: 'Unknown error' })
  })

  it('does not inspect or serialize arbitrary object properties', () => {
    const error = {
      secret: 'PRIVATE_DOCUMENT',
      get message() {
        throw new Error('Do not inspect')
      },
    }
    expect(serializeDocxKitError(error)).toEqual({ message: '[object Object]' })
  })

  it.each(Object.values(ERROR_CODES))(
    'links %s to the matching documentation page',
    code => {
      const error = new DocxKitError(code, 'Original message')
      expect(error.docs).toBe(
        `https://docx-kit.ntnyq.dev/errors/${DIAGNOSTIC_CODES[code].toLowerCase()}/`,
      )
      expect(serializeDocxKitError(error).docs).toBe(error.docs)
    },
  )
})
