import * as docxKit from 'docx-kit'
import { describe, expect, it, vi } from 'vitest'
import {
  CHECKLIST_CODE,
  COMPREHENSIVE_CODE,
  REFERENCES_CODE,
  RICH_CONTENT_CODE,
} from '../constants/templates'
import { executePlaygroundCode } from './playgroundRunner'
import { prepareCode } from './prepareCode'

const BUILTIN_PLUGIN_FACTORIES = {
  badge: docxKit.badgePlugin,
  barcode: docxKit.barcodePlugin,
  bibliography: docxKit.bibliographyPlugin,
  callout: docxKit.calloutPlugin,
  changelog: docxKit.changelogPlugin,
  checklist: docxKit.checklistPlugin,
  codeBlock: docxKit.codeBlockPlugin,
  coverPage: docxKit.coverPagePlugin,
  crossReference: docxKit.crossReferencePlugin,
  dataTable: docxKit.dataTablePlugin,
  divider: docxKit.dividerPlugin,
  echarts: docxKit.echartsPlugin,
  figure: docxKit.figurePlugin,
  glossary: docxKit.glossaryPlugin,
  invoice: docxKit.invoicePlugin,
  letterhead: docxKit.letterheadPlugin,
  meetingMinutes: docxKit.meetingMinutesPlugin,
  metricCard: docxKit.metricCardPlugin,
  pageNumber: docxKit.pageNumberPlugin,
  propertyTable: docxKit.propertyTablePlugin,
  qrcode: docxKit.qrcodePlugin,
  signatureBlock: docxKit.signatureBlockPlugin,
  timeline: docxKit.timelinePlugin,
  toc: docxKit.tocPlugin,
  watermark: docxKit.watermarkPlugin,
} as const

describe('playground runtime', () => {
  it('retries runtime initialization after a transient loading failure', async () => {
    const failure = new Error('Runtime temporarily unavailable')
    const combine = vi.spyOn(Promise, 'all').mockRejectedValueOnce(failure)
    const rejected = executePlaygroundCode('42')
    combine.mockRestore()
    await expect(rejected).rejects.toBe(failure)
    await expect(executePlaygroundCode('42')).resolves.toBe(42)
  })

  it('supports namespace imports for both runtime packages', async () => {
    await expect(
      executePlaygroundCode(`
      import * as kit from 'docx-kit'
      import * as native from 'docx'
      kit.createDocx().p(native.HeadingLevel.HEADING_1).toBlob()
    `),
    ).resolves.toBeInstanceOf(Blob)
  })

  it('allows imported bindings to share the executor namespace names', async () => {
    await expect(
      executePlaygroundCode(`
      import { createDocx as docxKit } from 'docx-kit'
      import { HeadingLevel as docx } from 'docx'
      docxKit().p(docx.HEADING_1).toBlob()
    `),
    ).resolves.toBeInstanceOf(Blob)
  })

  it('preserves import examples inside strings and comments', async () => {
    const sample =
      "\nimport { createDocx } from 'docx-kit'\nimport 'external-package'\n"
    await expect(
      executePlaygroundCode(`
      import { createDocx } from 'docx-kit'
      /*
      import { notAnImport } from 'docx-kit'
      */
      const sample = ${'`'}${sample}${'`'}
      sample
    `),
    ).resolves.toBe(sample)
  })

  it.each([
    `createDocx()\n  .p('Hello')\n  .toBlob()`,
    `createDocx({\n  page: { margin: 0 }\n}).toBlob() // final comment`,
    `await createDocx()\n  .p('Hello')\n  .toBlob()\n/* trailing comment */`,
  ])('returns multiline expressions', async expression => {
    await expect(
      executePlaygroundCode(
        `import { createDocx } from 'docx-kit'\n${expression}`,
      ),
    ).resolves.toBeInstanceOf(Blob)
  })
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

    expect(source).toContain(
      'const { createDocx: makeDocx, span } = arguments[0];',
    )
    expect(source).toContain(
      'const { Paragraph: DocxParagraph } = arguments[1];',
    )
    expect(source).not.toContain('DocxKitConfig')
    expect(source).not.toContain('external-package')
  })

  it.each([
    'return await Promise.resolve(42)',
    'if (true) return await Promise.resolve(42)',
    'const value: number = 42; value as number',
    'const value = 42; value\ntype Result = number',
    'const value = 42; value\ninterface Result { value: number }',
    'const value = 42; value\nimport type { Result } from "external-package"',
    'const value = 42; value\nimport "external-package"',
    'enum Status { Ready = 42 }; Status.Ready',
    'class Value { constructor(public value: number) {} }; new Value(42).value',
    'class Value { value: number = 42 }; new Value().value',
    'const value: { nested?: { count: number } } = {}; value.nested?.count ?? 42',
  ])('preserves TypeScript runtime behavior for %s', async source => {
    await expect(executePlaygroundCode(source)).resolves.toBe(42)
  })

  it.each([
    '42; import { createDocx } from "docx-kit"',
    '42; import * as kit from "docx-kit"',
    '42; const following = 1',
    '42;;',
  ])(
    'does not auto-return an expression before a following statement: %s',
    async source => {
      await expect(executePlaygroundCode(source)).resolves.toBeUndefined()
    },
  )

  it('keeps string import names and UTF-16 source ranges intact', async () => {
    await expect(
      executePlaygroundCode(`
        const label = '文档📄'
        import { "HeadingLevel" as headings } from 'docx'
        [label, headings.HEADING_1]
      `),
    ).resolves.toEqual(['文档📄', 'Heading1'])
  })

  it('preserves dynamic imports', () => {
    expect(prepareCode('import("external-package")')).toContain(
      'import("external-package")',
    )
  })

  it.each(['const value: number =', 'return 42; const value ='])(
    'rejects invalid code: %s',
    source => {
      expect(() => prepareCode(source)).toThrow(SyntaxError)
    },
  )

  it('transforms createDocx configuration through runtime overrides', async () => {
    const transformConfig = vi.fn(config => config)

    await expect(
      executePlaygroundCode(
        `
          import { createDocx } from 'docx-kit'

          const doc = createDocx({ metadata: { title: 'Lazy runtime' } })
          doc.toBlob()
        `,
        { transformConfig },
      ),
    ).resolves.toBeInstanceOf(Blob)

    expect(transformConfig).toHaveBeenCalledWith({
      metadata: { title: 'Lazy runtime' },
    })
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

  it('runs the reference and glossary example', async () => {
    await expect(
      executePlaygroundCode(REFERENCES_CODE),
    ).resolves.toBeInstanceOf(Blob)
  })

  it('runs the complete checklist example', async () => {
    await expect(executePlaygroundCode(CHECKLIST_CODE)).resolves.toBeInstanceOf(
      Blob,
    )
  })
})
