import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { createDocx } from '../../src/builder/createDocx'
import { resolveStyle } from '../../src/style/normalizeStyle'

it('lets inline shorthand reset longhands from defaults, inheritance and later classes', async () => {
  const document = createDocx({
    defaults: { paragraph: { marginLeft: 24, marginTop: 12 } },
    styles: {
      base: { marginRight: 24 },
      body: { extends: 'base', marginBottom: 24 },
    },
  }).p('Reset', { className: 'body', style: { margin: 0 } })
  const zip = await JSZip.loadAsync(await document.toBuffer())
  const xml = await zip.file('word/document.xml')?.async('string')
  expect(xml).toContain('<w:spacing w:after="0" w:before="0"/>')
  expect(xml).toContain('<w:ind w:left="0" w:right="0"/>')
})

it('resets inherited longhands with a token shorthand while keeping same-layer longhands', () => {
  const style = resolveStyle({
    className: ['child', 'last'],
    theme: { spacing: { small: 2 } },
    styles: {
      base: { marginLeft: 24, marginTop: 24 },
      child: { extends: 'base', margin: '$spacing.small', marginLeft: 3 },
      last: { marginRight: 4 },
    },
  })
  expect(style).toEqual({ margin: 2, marginLeft: 3, marginRight: 4 })
})

it('lets later aliases override earlier layers without changing same-layer precedence', () => {
  expect(
    resolveStyle({
      base: { bold: false, italic: false, letterSpacing: 2 },
      inline: { characterSpacing: 0, fontStyle: 'italic', fontWeight: 'bold' },
    }),
  ).toEqual({ characterSpacing: 0, fontStyle: 'italic', fontWeight: 'bold' })
  expect(resolveStyle({ inline: { bold: false, fontWeight: 'bold' } })).toEqual(
    { bold: false, fontWeight: 'bold' },
  )
})

it('preserves default-layer precedence through rich text, links and bookmarks', async () => {
  const document = createDocx({
    defaults: {
      paragraph: { fontStyle: 'italic', fontWeight: 'bold', margin: 0 },
      text: { bold: false, italic: false, marginLeft: 24 },
    },
  })
    .p('Plain')
    .p([
      { text: 'Rich', type: 'text' },
      { children: ['Link'], type: 'hyperlink', url: 'https://example.com' },
      { children: ['Bookmark'], name: 'target', type: 'bookmark' },
    ])
  const zip = await JSZip.loadAsync(await document.toBuffer())
  const xml = await zip.file('word/document.xml')?.async('string')
  expect(xml?.match(/<w:b\/>/g)).toHaveLength(4)
  expect(xml?.match(/<w:i\/>/g)).toHaveLength(4)
  expect(xml?.match(/<w:ind w:left="0" w:right="0"\/>/g)).toHaveLength(2)
})
