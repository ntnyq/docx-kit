import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

describe('reference plugins public entry points', () => {
  it.each([
    ['browser source', () => import('../src/browser')],
    ['node source', () => import('../src/node')],
    ['browser dist', () => import('../dist/browser.js')],
    ['node dist', () => import('../dist/node.js')],
  ] as const)(
    'exports fluent and JSON documents through %s',
    async (_, load) => {
      const mod = await load()
      const entries = [{ id: 'source', title: 'Source document' }]
      const fluent = mod
        .createDocx()
        .use(mod.crossReferencePlugin())
        .use(mod.bibliographyPlugin({ entries }))
        .use(mod.glossaryPlugin())
        .bookmark('overview', ['Overview'])
        .plugin('crossReference', { fallback: 'Overview', target: 'overview' })
        .plugin('bibliography', { ids: ['source'], mode: 'citation' })
        .plugin('bibliography', {})
        .plugin('glossary', {
          entries: [{ definition: 'Document format', term: 'DOCX' }],
        })
      const schema = await mod.renderDocx({
        content: [
          { children: ['Overview'], name: 'overview', type: 'bookmark' },
          {
            name: 'crossReference',
            options: { target: 'overview' },
            type: 'plugin',
          },
          { name: 'bibliography', options: { entries }, type: 'plugin' },
          {
            name: 'glossary',
            type: 'plugin',
            options: {
              entries: [{ definition: 'Document format', term: 'DOCX' }],
            },
          },
        ],
        plugins: [
          { plugin: mod.crossReferencePlugin(), type: 'inline' },
          { plugin: mod.bibliographyPlugin(), type: 'inline' },
          { plugin: mod.glossaryPlugin(), type: 'inline' },
        ],
      })
      for (const document of [fluent, schema]) {
        const archive = await JSZip.loadAsync(await document.toBuffer())
        const xml = await archive.file('word/document.xml')!.async('string')
        expect(xml).toContain('>REF overview \\h</w:instrText>')
        expect(xml).toContain('[1] Source document')
        expect(xml).toContain('Document format')
      }
    },
  )
})
