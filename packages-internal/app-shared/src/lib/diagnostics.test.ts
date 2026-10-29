import { createDocxKitError, DocxKitError } from '@docxkit/errors'
import { describe, expect, it } from 'vitest'
import { getDiagnosticDocsUrl, serializeDocxKitError } from './diagnostics'

describe('localized diagnostic guidance', () => {
  it.each([
    ['en-US', ''],
    ['zh-CN', '/zh-CN'],
    ['zh-TW', '/zh-TW'],
    ['fr', ''],
  ])('links %s to its error page', (locale, prefix) => {
    const error = createDocxKitError('PLUGIN_NOT_REGISTERED', {
      pluginName: 'qrcode',
    })
    expect(getDiagnosticDocsUrl(serializeDocxKitError(error), locale)).toBe(
      `https://docx-kit.ntnyq.dev${prefix}/errors/docxkit_plugin_not_registered/`,
    )
  })

  it('leaves unknown errors readable without inventing a repair link', () => {
    const error = serializeDocxKitError(
      new DocxKitError('CUSTOM', 'Custom failure'),
    )
    expect(error.message).toBe('Custom failure')
    expect(getDiagnosticDocsUrl(error, 'zh-CN')).toBeUndefined()
  })
})
