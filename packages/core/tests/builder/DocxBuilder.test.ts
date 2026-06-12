import { describe, expect, it } from 'vitest'
import { DocxBuilder } from '../../src/builder/DocxBuilder'

describe('DocxBuilder', () => {
  describe('constructor & toJSON', () => {
    it('creates builder with no config', () => {
      const builder = new DocxBuilder()
      expect(builder).toBeInstanceOf(DocxBuilder)
    })

    it('creates builder with config', () => {
      const builder = new DocxBuilder({
        metadata: { title: 'Test' },
        page: { size: 'A4' },
      })
      expect(builder).toBeInstanceOf(DocxBuilder)
    })

    it('toJSON returns config + content', () => {
      const builder = new DocxBuilder({ page: { size: 'A4' } })
      builder.h1('Title')
      const json = builder.toJSON()
      expect(json.content).toHaveLength(1)
      expect(json.content![0]).toMatchObject({ text: 'Title', type: 'heading' })
    })
  })

  describe('headings', () => {
    it('h1 adds a level-1 heading node', () => {
      const builder = new DocxBuilder()
      builder.h1('Main Title')
      const json = builder.toJSON()
      expect(json.content![0]).toMatchObject({
        level: 1,
        text: 'Main Title',
        type: 'heading',
      })
    })

    it('h2 adds a level-2 heading', () => {
      const builder = new DocxBuilder()
      builder.h2('Section')
      const json = builder.toJSON()
      expect(json.content![0]).toMatchObject({
        level: 2,
        type: 'heading',
      })
    })

    it('h3 adds a level-3 heading', () => {
      const builder = new DocxBuilder()
      builder.h3('Subsection')
      const json = builder.toJSON()
      expect(json.content![0]).toMatchObject({ level: 3, type: 'heading' })
    })

    it('h4 adds a level-4 heading', () => {
      const builder = new DocxBuilder()
      builder.h4('Deep')
      expect(builder.toJSON().content![0]).toMatchObject({
        level: 4,
        type: 'heading',
      })
    })

    it('h5 adds a level-5 heading', () => {
      const builder = new DocxBuilder()
      builder.h5('Deeper')
      expect(builder.toJSON().content![0]).toMatchObject({
        level: 5,
        type: 'heading',
      })
    })

    it('h6 adds a level-6 heading', () => {
      const builder = new DocxBuilder()
      builder.h6('Deepest')
      expect(builder.toJSON().content![0]).toMatchObject({
        level: 6,
        type: 'heading',
      })
    })

    it('heading accepts className and style overrides', () => {
      const builder = new DocxBuilder()
      builder.h1('Titled', { className: 'title', style: { color: '#f00' } })
      const node = builder.toJSON().content![0]
      expect(node).toMatchObject({
        className: 'title',
        style: { color: '#f00' },
      })
    })
  })

  describe('paragraph', () => {
    it('p adds a paragraph node', () => {
      const builder = new DocxBuilder()
      builder.p('Hello world')
      expect(builder.toJSON().content![0]).toMatchObject({
        text: 'Hello world',
        type: 'paragraph',
      })
    })

    it('p accepts className and style', () => {
      const builder = new DocxBuilder()
      builder.p('Styled', { className: 'body', style: { fontSize: '14pt' } })
      const node = builder.toJSON().content![0]
      expect(node).toMatchObject({
        className: 'body',
        style: { fontSize: '14pt' },
      })
    })
  })

  describe('pageBreak', () => {
    it('adds a pageBreak node', () => {
      const builder = new DocxBuilder()
      builder.pageBreak()
      expect(builder.toJSON().content![0]).toMatchObject({
        type: 'pageBreak',
      })
    })
  })

  describe('image', () => {
    it('adds an image node', () => {
      const builder = new DocxBuilder()
      builder.image({ data: new Uint8Array([1, 2, 3]), height: 50, width: 100 })
      const node = builder.toJSON().content![0]
      expect(node).toMatchObject({
        height: 50,
        type: 'image',
        width: 100,
      })
    })
  })

  describe('table', () => {
    it('adds a table node', () => {
      const builder = new DocxBuilder()
      builder.table({
        columns: [{ key: 'name', title: 'Name' }],
        data: [{ name: 'Alice' }],
      })
      const node = builder.toJSON().content![0]
      expect(node).toMatchObject({ type: 'table' })
    })
  })

  describe('plugin', () => {
    it('adds a plugin node', () => {
      const builder = new DocxBuilder() as any
      builder.plugins.toMap().set('custom', { name: 'custom' })
      builder.plugin('custom', { value: 42 })
      const node = builder.toJSON().content![0]
      expect(node).toMatchObject({ name: 'custom', type: 'plugin' })
    })

    it('plugin accepts optional style param', () => {
      const builder = new DocxBuilder() as any
      builder.plugins.toMap().set('custom', { name: 'custom' })
      builder.plugin('custom', {}, { color: '#f00' })
      const node = builder.toJSON().content![0]
      expect(node).toMatchObject({ name: 'custom', style: { color: '#f00' } })
    })
  })

  describe('use', () => {
    it('registers a plugin and returns builder', () => {
      const builder = new DocxBuilder()
      const result = builder.use({ name: 'myPlugin', render: () => 'ok' })
      expect(result).toBe(builder)
      expect((builder as any).plugins.toMap().has('myPlugin')).toBe(true)
    })
  })

  describe('add', () => {
    it('adds raw node and returns this', () => {
      const builder = new DocxBuilder()
      const result = builder.add({ type: 'pageBreak' })
      expect(result).toBe(builder)
      expect(builder.toJSON().content).toHaveLength(1)
    })
  })

  describe('chaining', () => {
    it('supports fluent chaining of multiple node types', () => {
      const builder = new DocxBuilder()
      builder
        .h1('Title', { className: 'title' })
        .p('Intro')
        .pageBreak()
        .h2('Section')
        .p('Content')

      const json = builder.toJSON()
      expect(json.content).toHaveLength(5)
      expect(json.content![0]).toMatchObject({
        className: 'title',
        level: 1,
        type: 'heading',
      })
      expect(json.content![2]).toMatchObject({ type: 'pageBreak' })
    })
  })
})
