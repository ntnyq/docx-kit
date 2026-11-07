import { expectTypeOf } from 'vitest'
import * as Browser from '../src/browser'
import type { PluginSource } from '@docxkit/core'
import type * as Node from '../src/node'

expectTypeOf<Node.CrossReferenceOptions>().toEqualTypeOf<Browser.CrossReferenceOptions>()
expectTypeOf<Node.BibliographyOptions>().toEqualTypeOf<Browser.BibliographyOptions>()
expectTypeOf<Node.GlossaryOptions>().toEqualTypeOf<Browser.GlossaryOptions>()

const doc = Browser.createDocx()
  .use(Browser.crossReferencePlugin())
  .use(Browser.bibliographyPlugin())
  .use(Browser.glossaryPlugin())

doc.plugin('crossReference', { kind: 'number', target: 'chapter' })
doc.plugin('bibliography', { ids: ['source'], mode: 'citation' })
doc.plugin('bibliography', { entries: [] })
doc.plugin('glossary', { entries: [], sort: 'none' })

// @ts-expect-error A citation requires source IDs.
doc.plugin('bibliography', { mode: 'citation' })
// @ts-expect-error Invalid cross-reference kind.
doc.plugin('crossReference', { kind: 'unknown', target: 'chapter' })
// @ts-expect-error A glossary entry requires a definition.
doc.plugin('glossary', { entries: [{ term: 'API' }] })
doc.plugin('crossReference', {
  // @ts-expect-error Paragraph alignment belongs in the paragraph part.
  styles: { text: { textAlign: 'right' } },
  target: 'chapter',
})

// Inline schema sources accept specifically typed render options without casts.
const sources: PluginSource[] = [
  { plugin: Browser.crossReferencePlugin(), type: 'inline' },
  { plugin: Browser.bibliographyPlugin(), type: 'inline' },
  { plugin: Browser.glossaryPlugin(), type: 'inline' },
]
Browser.renderDocx({ content: [], plugins: sources })
