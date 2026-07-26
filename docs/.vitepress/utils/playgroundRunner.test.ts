import * as docxKit from 'docx-kit'
import { describe, expect, it } from 'vitest'
import { COMPREHENSIVE_CODE, RICH_CONTENT_CODE } from '../constants/templates'
import { executePlaygroundCode } from './playgroundRunner'
import { prepareCode } from './prepareCode'

const BUILTIN_PLUGIN_FACTORIES = {
  badge: docxKit.badgePlugin,
  barcode: docxKit.barcodePlugin,
  callout: docxKit.calloutPlugin,
  changelog: docxKit.changelogPlugin,
  codeBlock: docxKit.codeBlockPlugin,
  coverPage: docxKit.coverPagePlugin,
  dataTable: docxKit.dataTablePlugin,
  divider: docxKit.dividerPlugin,
  echarts: docxKit.echartsPlugin,
  invoice: docxKit.invoicePlugin,
  letterhead: docxKit.letterheadPlugin,
  meetingMinutes: docxKit.meetingMinutesPlugin,
  pageNumber: docxKit.pageNumberPlugin,
  propertyTable: docxKit.propertyTablePlugin,
  qrcode: docxKit.qrcodePlugin,
  signatureBlock: docxKit.signatureBlockPlugin,
  timeline: docxKit.timelinePlugin,
  toc: docxKit.tocPlugin,
  watermark: docxKit.watermarkPlugin,
} as const

describe('playground runtime', () => {
  it('exposes all built-in plugins as factories with matching plugin names', () => {
    for (const [name, createPlugin] of Object.entries(
      BUILTIN_PLUGIN_FACTORIES,
    )) {
      expect(createPlugin).toBeTypeOf('function')
      expect(createPlugin()).toMatchObject({
        name,
        render: expect.any(Function),
      })
    }
  })

  it('maps runtime imports, aliases, and type-only imports correctly', () => {
    const source = prepareCode(`
      import {
        createDocx as makeDocx,
        type DocxKitConfig,
        span,
      } from 'docx-kit'
      import { Paragraph as DocxParagraph } from 'docx'
      import type { ExternalOptions } from 'external-package'

      const config: DocxKitConfig & Partial<ExternalOptions> = {}
      const doc = makeDocx(config)
      new DocxParagraph({ children: [], text: span('ready').text })
      doc.toBlob()
    `)

    expect(source).toContain('const { createDocx: makeDocx, span } = docxKit;')
    expect(source).toContain('const { Paragraph: DocxParagraph } = docx;')
    expect(source).not.toContain('DocxKitConfig')
    expect(source).not.toContain('external-package')
  })

  it('runs the custom-plugin comprehensive preset', async () => {
    await expect(
      executePlaygroundCode(COMPREHENSIVE_CODE),
    ).resolves.toBeInstanceOf(Blob)
  }, 20_000)

  it('runs the rich-content preset with a partial stylesheet', async () => {
    await expect(
      executePlaygroundCode(RICH_CONTENT_CODE),
    ).resolves.toBeInstanceOf(Blob)
  }, 20_000)
})
