import { Lexer, marked } from 'marked'
import { describe, expect, it } from 'vitest'
import { markdownToNodes } from '../src'

describe('markdownToNodes', () => {
  it('converts headings, paragraphs and thematic breaks to the existing DSL', () => {
    expect(
      markdownToNodes(
        '# **Title** &amp; [link](https://example.com)\n\nBody\n\n---',
      ),
    ).toEqual([
      { level: 1, text: 'Title & link', type: 'heading' },
      { children: [{ text: 'Body', type: 'text' }], type: 'paragraph' },
      { type: 'thematicBreak' },
    ])
    expect(markdownToNodes('')).toEqual([])
    expect(markdownToNodes(' \n\n')).toEqual([])
    expect(markdownToNodes('Title\n=====')[0]).toMatchObject({
      level: 1,
      text: 'Title',
    })
    expect(markdownToNodes('###### Last')[0]).toMatchObject({
      level: 6,
      text: 'Last',
    })
  })

  it('preserves nested emphasis, strikethrough, literal code and escapes', () => {
    expect(
      markdownToNodes('**bold *both*** ~~gone~~ `&amp;` \\* &copy;')[0],
    ).toEqual({
      type: 'paragraph',
      children: [
        { style: { bold: true }, text: 'bold ', type: 'text' },
        { style: { bold: true, italic: true }, text: 'both', type: 'text' },
        { text: ' ', type: 'text' },
        { style: { strike: true }, text: 'gone', type: 'text' },
        { text: ' ', type: 'text' },
        { style: { fontFamily: 'Consolas' }, text: '&amp;', type: 'text' },
        { text: ' ', type: 'text' },
        { text: '*', type: 'text' },
        { text: ' ©', type: 'text' },
      ],
    })
  })

  it('keeps soft breaks as spaces and hard breaks as Word run breaks', () => {
    expect(markdownToNodes('one\ntwo')[0]).toMatchObject({
      children: [{ text: 'one two' }],
    })
    for (const source of ['one  \ntwo', 'one\\\ntwo']) {
      expect(markdownToNodes(source)[0]).toMatchObject({
        children: [
          { text: 'one' },
          { style: { docx: { break: 1 } }, text: '' },
          { text: 'two' },
        ],
      })
    }
    expect(markdownToNodes('one\ntwo', { breaks: true })[0]).toMatchObject({
      children: [
        { text: 'one' },
        { style: { docx: { break: 1 } } },
        { text: 'two' },
      ],
    })
  })

  it('resolves references and preserves formatted link labels', () => {
    expect(
      markdownToNodes(
        '[**Docs**][ref]\n\n[ref]: https://example.com/?a=1&amp;b=2',
      )[0],
    ).toMatchObject({
      children: [
        {
          children: [{ style: { bold: true }, text: 'Docs', type: 'text' }],
          type: 'hyperlink',
          url: 'https://example.com/?a=1&b=2',
        },
      ],
    })
    expect(markdownToNodes('<https://example.com/?q=&amp;>')[0]).toMatchObject({
      children: [
        {
          children: [{ text: 'https://example.com/?q=&amp;' }],
          url: 'https://example.com/?q=&amp;',
        },
      ],
    })
  })

  it('passes image sources through and distinguishes block and inline images', () => {
    expect(
      markdownToNodes('![**Alt**](./image.png)', { imageSize: { width: 240 } }),
    ).toEqual([{ alt: 'Alt', data: './image.png', type: 'image', width: 240 }])
    expect(
      markdownToNodes('See ![Alt](https://example.com/image.png) here')[0],
    ).toMatchObject({
      type: 'paragraph',
      children: [
        { text: 'See ' },
        { alt: 'Alt', type: 'image' },
        { text: ' here' },
      ],
    })
    expect(
      markdownToNodes('[![Alt](image.png)](https://example.com)')[0],
    ).toMatchObject({
      children: [{ children: [{ text: 'Alt' }], type: 'hyperlink' }],
    })
  })

  it('keeps nested mixed lists in source order and resumes ordered parents', () => {
    expect(
      markdownToNodes('3. parent\n   - **child**\n   - next\n4. tail'),
    ).toEqual([
      {
        items: [{ children: [{ text: 'parent', type: 'text' }] }],
        level: 0,
        start: 3,
        type: 'numberedList',
      },
      {
        level: 1,
        type: 'bulletList',
        items: [
          {
            children: [{ style: { bold: true }, text: 'child', type: 'text' }],
          },
          { children: [{ text: 'next', type: 'text' }] },
        ],
      },
      {
        items: [{ children: [{ text: 'tail', type: 'text' }] }],
        level: 0,
        start: 4,
        type: 'numberedList',
      },
    ])
  })

  it('retains loose list continuation blocks and code blocks', () => {
    const nodes = markdownToNodes(
      '- first\n\n  second\n\n  ```js\n  const n = 1\n  ```\n\n- last',
    )
    expect(nodes.map(node => node.type)).toEqual([
      'bulletList',
      'paragraph',
      'plugin',
      'bulletList',
    ])
    expect(nodes[1]).toMatchObject({ children: [{ text: 'second' }] })
    expect(nodes[2]).toMatchObject({
      name: 'codeBlock',
      options: { code: 'const n = 1', language: 'js' },
    })
  })

  it('uses native checkbox inline nodes for task lists', () => {
    const [node] = markdownToNodes('- [x] **Done**\n- [ ] Todo')
    expect(node).toMatchObject({
      type: 'bulletList',
      items: [
        {
          children: [
            { checked: true, type: 'checkbox' },
            { text: ' ' },
            { style: { bold: true }, text: 'Done' },
          ],
        },
        {
          children: [
            { checked: false, type: 'checkbox' },
            { text: ' ' },
            { text: 'Todo' },
          ],
        },
      ],
    })
  })

  it('converts tables with alignment, readable data and rich cell renderers', () => {
    const [node] = markdownToNodes(
      '| **Name** | Count |\n| :--- | ---: |\n| **Alpha** | [two](https://example.com) |\n| Beta | |',
    )
    if (node.type !== 'table') {
      throw new Error('Expected a table')
    }
    expect(node.columns).toMatchObject([
      { align: 'left', title: 'Name' },
      { align: 'right', title: 'Count' },
    ])
    expect(node.data).toEqual([
      { column0: 'Alpha', column1: 'two' },
      { column0: 'Beta', column1: '' },
    ])
    expect(node.columns[0].render?.('Alpha', node.data[0], 0)).toEqual([
      { style: { bold: true }, text: 'Alpha', type: 'text' },
    ])
  })

  it.each(['filter', 'reverse'] as const)(
    'preserves the matching rich cells when table rows %s',
    operation => {
      const [node] = markdownToNodes(
        '| Name |\n| --- |\n| **Alpha** |\n| *Beta* |',
      )
      if (node.type !== 'table') {
        throw new Error('Expected a table')
      }
      node.data =
        operation === 'filter'
          ? node.data.filter(row => row.column0 === 'Beta')
          : node.data.toReversed()

      expect(
        node.columns[0].render?.(node.data[0].column0, node.data[0], 0),
      ).toEqual([{ style: { italic: true }, text: 'Beta', type: 'text' }])
    },
  )

  it('renders current values for edited cells, appended rows and copied rows', () => {
    const [node] = markdownToNodes(
      '| Name | Value |\n| --- | --- |\n| **Alpha** | *Ten* |',
    )
    if (node.type !== 'table') {
      throw new Error('Expected a table')
    }
    node.data[0].column0 = 'Updated'
    node.data.push(
      { column0: 'Added', column1: 0 },
      { ...node.data[0] },
      { column0: null },
    )

    expect(node.columns[0].render?.('Updated', node.data[0], 0)).toBe('Updated')
    expect(node.columns[1].render?.('Ten', node.data[0], 0)).toEqual([
      { style: { italic: true }, text: 'Ten', type: 'text' },
    ])
    expect(node.columns[0].render?.('Added', node.data[1], 1)).toBe('Added')
    expect(node.columns[1].render?.(0, node.data[1], 1)).toBe('0')
    expect(node.columns[0].render?.('Updated', node.data[2], 2)).toBe('Updated')
    expect(node.columns[0].render?.(null, node.data[3], 3)).toBe('')
    expect(node.columns[1].render?.(undefined, node.data[3], 3)).toBe('')
  })

  it('does not mix formatting between rows with identical text', () => {
    const [node] = markdownToNodes(
      '| Name |\n| --- |\n| **Same** |\n| *Same* |',
    )
    if (node.type !== 'table') {
      throw new Error('Expected a table')
    }
    node.data.reverse()

    expect(node.columns[0].render?.('Same', node.data[0], 0)).toEqual([
      { style: { italic: true }, text: 'Same', type: 'text' },
    ])
    expect(node.columns[0].render?.('Same', node.data[1], 1)).toEqual([
      { style: { bold: true }, text: 'Same', type: 'text' },
    ])
  })

  it('maps fenced and indented code blocks to the codeBlock plugin', () => {
    expect(
      markdownToNodes('```ts title=demo\nconst x = "&amp;"\n```')[0],
    ).toEqual({
      name: 'codeBlock',
      options: { code: 'const x = "&amp;"', language: 'ts' },
      type: 'plugin',
    })
    expect(markdownToNodes('    const x = 1')[0]).toMatchObject({
      name: 'codeBlock',
      options: { code: 'const x = 1' },
    })
  })

  it('attaches existing classes and preserves paragraph styling inside quotes', () => {
    expect(
      markdownToNodes('> Quote', {
        classNames: { blockquote: 'quote', paragraph: 'body' },
      })[0],
    ).toMatchObject({ className: ['body', 'quote'] })
    expect(markdownToNodes('> Quote')[0]).toMatchObject({
      style: { marginLeft: '12pt' },
    })
    expect(markdownToNodes('> # Title')[0]).toEqual({
      level: 1,
      style: { marginLeft: '12pt' },
      text: 'Title',
      type: 'heading',
    })
    expect(
      markdownToNodes('`code`', { classNames: { inlineCode: 'mono' } })[0],
    ).toMatchObject({ children: [{ className: 'mono', text: 'code' }] })
  })

  it.each([
    { blockquote: ' quote\nmuted ', paragraph: ' body\tlead ' },
    { blockquote: ['quote', 'muted'], paragraph: ['body', 'lead'] },
    { blockquote: ['quote\tmuted'], paragraph: [' body lead '] },
  ])('normalizes quote class strings and arrays: %j', classNames => {
    expect(markdownToNodes('> Quote', { classNames })[0]).toMatchObject({
      className: ['body', 'lead', 'quote', 'muted'],
    })
  })

  it('normalizes multiple quote classes without a paragraph class', () => {
    expect(
      markdownToNodes('> Quote', {
        classNames: { blockquote: 'quote muted' },
      })[0],
    ).toMatchObject({ className: ['quote', 'muted'] })
  })

  it('lets transforms reuse custom plugins and omit blocks', () => {
    const nodes = markdownToNodes('> Note\n\n---', {
      transform(nodes, context) {
        if (context.type === 'blockquote') {
          return [
            {
              name: 'callout',
              options: { content: context.raw, type: 'info' },
              type: 'plugin',
            },
          ]
        }
        return context.type === 'hr' ? [] : nodes
      },
    })
    expect(nodes).toEqual([
      {
        name: 'callout',
        options: { content: '> Note', type: 'info' },
        type: 'plugin',
      },
    ])
  })

  it('preserves raw HTML literally and supports disabling GFM', () => {
    expect(markdownToNodes('<script>alert(1)</script>')[0]).toEqual({
      text: '<script>alert(1)</script>',
      type: 'paragraph',
    })
    expect(markdownToNodes('~~text~~', { gfm: false })[0]).toMatchObject({
      children: [{ text: '~~text~~' }],
    })
  })

  it('does not share global parser extensions or per-call options', () => {
    const defaults = { ...marked.defaults }
    try {
      marked.use({ gfm: false })
      expect(markdownToNodes('~~text~~')[0]).toMatchObject({
        children: [{ style: { strike: true } }],
      })
      markdownToNodes('a\nb', { breaks: true })
      expect(markdownToNodes('a\nb')[0]).toMatchObject({
        children: [{ text: 'a b' }],
      })
      expect(new Lexer().lex('text')).toHaveLength(1)
    } finally {
      marked.setOptions(defaults)
    }
  })
})
