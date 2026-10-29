import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'
import type {
  BlockNode,
  DocxKitConfig,
  DocxPlugin,
  RevisionNode,
} from '@docxkit/types'

const revision: RevisionNode = {
  author: 'Reviewer',
  children: ['Inherited'],
  date: '2026-10-30T00:00:00Z',
  revisionId: 1,
  type: 'insertedText',
}

async function renderPackage(
  nodes: BlockNode[],
  config: DocxKitConfig = {},
  plugins = new Map<string, DocxPlugin>(),
) {
  const document = await compileDocument({ config, nodes, plugins })
  return JSZip.loadAsync(await Packer.toBuffer(document))
}

describe('revision and checkbox contracts', () => {
  it.each(['insertedText', 'deletedText'] as const)(
    'inherits paragraph styles in %s and checkbox labels while preserving overrides',
    async type => {
      const archive = await renderPackage([
        {
          style: { color: '#FF0000', fontSize: 20 },
          type: 'paragraph',
          children: [
            {
              ...revision,
              type,
              children: [
                'Inherited',
                {
                  style: { color: '#0000FF', fontSize: 10 },
                  text: 'Override',
                  type: 'text',
                },
              ],
            },
            { label: 'Checkbox', type: 'checkbox' },
          ],
        },
      ])
      const xml = await archive.file('word/document.xml')!.async('string')
      expect(xml.match(/<w:color w:val="FF0000"\/>/g)).toHaveLength(2)
      expect(xml.match(/<w:sz w:val="40"\/>/g)).toHaveLength(2)
      expect(xml).toContain('<w:color w:val="0000FF"/>')
      expect(xml).toContain('<w:sz w:val="20"/>')
    },
  )

  it.each<BlockNode>([
    { items: [{ children: [revision] }], type: 'bulletList' },
    {
      author: 'Reviewer',
      children: [revision],
      comment: ['Body'],
      type: 'comment',
    },
    {
      content: [{ children: [revision], type: 'paragraph' }],
      type: 'footnote',
    },
    {
      config: { header: { default: { children: [revision] } } },
      type: 'sectionBreak',
    },
    {
      columns: [{ key: 'value', title: 'Value', render: () => [revision] }],
      data: [{ value: 'x' }],
      type: 'table',
    },
  ])('enables tracking for revisions compiled inside $type', async node => {
    const archive = await renderPackage([node])
    expect(await archive.file('word/settings.xml')!.async('string')).toContain(
      '<w:trackRevisions/>',
    )
  })

  it('tracks revisions emitted by plugins and honors an explicit false setting', async () => {
    const plugins = new Map<string, DocxPlugin>([
      [
        'revision',
        {
          name: 'revision',
          render: (_options, context) => context.compileNode(revision),
        },
      ],
    ])
    const nodes: BlockNode[] = [
      { name: 'revision', options: {}, type: 'plugin' },
    ]
    for (const enabled of [true, false]) {
      const archive = await renderPackage(
        nodes,
        enabled ? {} : { features: { trackRevisions: false } },
        plugins,
      )
      const xml = await archive.file('word/settings.xml')!.async('string')
      expect(xml).toContain(
        enabled ? '<w:trackRevisions/>' : '<w:trackRevisions w:val="false"/>',
      )
    }
  })
})
