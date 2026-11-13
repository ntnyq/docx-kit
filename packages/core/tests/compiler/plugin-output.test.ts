import { Paragraph, TableOfContents } from 'docx'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { createDocx } from '../../src/builder/createDocx'
import { definePlugin } from '../../src/builder/definePlugin'
import { span } from '../../src/dsl/helpers'

it.each([
  null,
  undefined,
  'raw-string',
  span('DSL text'),
  [new Paragraph('Valid'), 'invalid'],
])('rejects unsupported plugin output %j during compilation', async output => {
  const document = createDocx()
    .use(definePlugin({ name: 'invalid', render: () => output }))
    .plugin('invalid', {})
  await expect(document.toDocument()).rejects.toMatchObject({
    cause: expect.any(TypeError),
    code: 'PLUGIN_RENDER_FAILED',
    message: expect.stringContaining('invalid'),
  })
})

it('packs valid blocks and nested compiled DSL content in order', async () => {
  const document = createDocx()
    .use(
      definePlugin({
        name: 'valid',
        async render(_options, context) {
          return [
            new Paragraph('First'),
            await context.compileNode({ text: 'Second', type: 'paragraph' }),
          ]
        },
      }),
    )
    .plugin('valid', {})
  const zip = await JSZip.loadAsync(await document.toBuffer())
  const xml = await zip.file('word/document.xml')?.async('string')
  expect(xml).toMatch(/<w:p>.*First.*<\/w:p><w:p>.*Second.*<\/w:p>/)
})

it('supports empty output and table-of-contents blocks', async () => {
  const document = createDocx()
    .use(definePlugin({ name: 'empty', render: () => [] }))
    .use(
      definePlugin({
        name: 'toc',
        render: () => new TableOfContents('Contents'),
      }),
    )
    .plugin('empty', {})
    .plugin('toc', {})
  await expect(document.toBuffer()).resolves.toBeInstanceOf(Uint8Array)
})
