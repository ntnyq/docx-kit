import { unindent } from '@ntnyq/utils'

/**
 * Research report example shared by the Playground and plugin documentation.
 */
export const REFERENCES_CODE = unindent(`
  // #region references-example
  import {
    bibliographyPlugin,
    crossReferencePlugin,
    createDocx,
    glossaryPlugin,
    useTheme,
  } from 'docx-kit'

  const entries = [
    { id: 'architecture', authors: ['Alice Chen'], title: 'Document Architecture', year: '2026', publisher: 'Example Press' },
    { id: 'api', authors: ['Bob Lin'], title: 'API Design Notes', year: '2025', url: 'https://example.com/api' },
  ]

  const doc = createDocx({ theme: useTheme('ocean') })
    .use(crossReferencePlugin())
    .use(bibliographyPlugin({ entries }))
    .use(glossaryPlugin())
    .h1('Research Report')
    .bookmark('overview', ['Architecture overview'])
    .p('The report describes the document generation interface.')
    .plugin('bibliography', { mode: 'citation', ids: ['architecture', 'api', 'architecture'], prefix: 'Based on ', suffix: '.' })
    .plugin('crossReference', { target: 'overview', prefix: 'See ', fallback: 'Architecture overview', suffix: '.' })
    .plugin('crossReference', { target: 'overview', kind: 'page', prefix: 'Overview is on page ', fallback: '1', suffix: '.' })
    .plugin('glossary', {
      groupBy: 'group',
      entries: [
        { term: 'DSL', fullForm: 'Domain Specific Language', definition: 'A language designed for a particular domain.', group: 'Software' },
        { term: 'API', fullForm: 'Application Programming Interface', definition: 'The contract used by document-generation clients.', group: 'Software' },
        { term: 'OOXML', fullForm: 'Office Open XML', definition: 'The XML format used inside DOCX files.', group: 'Formats' },
      ],
    })
    .plugin('bibliography', {})

  doc.toBlob()
  // #endregion references-example
`)
