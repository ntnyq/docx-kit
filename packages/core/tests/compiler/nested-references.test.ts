import { Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { compileDocument } from '../../src/compiler/compileDocument'
import type { BlockNode } from '@docxkit/types'

describe('nested annotation definitions', () => {
  it('drains annotations discovered in alternating reference bodies', async () => {
    const nodes: BlockNode[] = [
      {
        author: 'Outer reviewer',
        children: [{ text: 'Body', type: 'text' }],
        type: 'comment',
        comment: [
          {
            type: 'paragraph',
            children: [
              {
                type: 'footnote',
                content: [
                  {
                    type: 'paragraph',
                    children: [
                      {
                        author: 'Inner reviewer',
                        children: [{ text: 'Nested note', type: 'text' }],
                        comment: ['Inner comment body'],
                        type: 'comment',
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    ]
    const document = await compileDocument({
      config: {},
      nodes,
      plugins: new Map(),
    })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    expect(await archive.file('word/comments.xml')?.async('string')).toContain(
      'Inner comment body',
    )
    expect(await archive.file('word/footnotes.xml')?.async('string')).toContain(
      'w:commentReference w:id="1"',
    )
  })
  it('exports comments registered while compiling footnote bodies', async () => {
    const nodes: BlockNode[] = [
      {
        type: 'footnote',
        content: [
          {
            type: 'paragraph',
            children: [
              {
                author: 'Reviewer',
                children: [{ text: 'Annotated note', type: 'text' }],
                comment: ['Footnote comment body'],
                type: 'comment',
              },
            ],
          },
        ],
      },
    ]
    const document = await compileDocument({
      config: {},
      nodes,
      plugins: new Map(),
    })
    const archive = await JSZip.loadAsync(await Packer.toBuffer(document))
    expect(await archive.file('word/footnotes.xml')?.async('string')).toContain(
      'w:commentReference w:id="0"',
    )
    expect(await archive.file('word/comments.xml')?.async('string')).toContain(
      'Footnote comment body',
    )
  })
})
