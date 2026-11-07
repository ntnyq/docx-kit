import {
  BUILTIN_PLUGIN_OPTION_SCHEMAS,
  generateSchemaGuide,
  generateToolDefinitions,
} from '@docxkit/ai'
import { renderDocx } from '@docxkit/core'
import Ajv from 'ajv'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { createBuiltinPluginSources } from '../src/plugins/catalog'
import { docxSchemaResource } from '../src/resources/schema'
import { buildPluginHelp } from '../src/tools/getPluginHelp'
import { validateSchema } from '../src/tools/validateSchema'

const validateResource = new Ajv({ strict: false }).compile(
  docxSchemaResource.schema,
)
const aiSchema = generateToolDefinitions().find(
  tool => tool.function.name === 'create_document',
)?.function.parameters
const validateAi = new Ajv({ strict: false }).compile(aiSchema ?? {})
const callout = { content: 'Message', title: 'Note', type: 'info' }
const table = { data: [{ score: 1 }, { score: 2 }], striped: true }
const node = (name: string, options: unknown) => ({
  name,
  options,
  type: 'plugin',
})

describe('plugin part style contracts', () => {
  it.each([
    node('callout', {
      ...callout,
      styles: {
        cell: { borderLeft: { color: '$colors.primary' }, margin: 0 },
        title: { bold: false, fontSize: '12pt' },
      },
    }),
    node('dataTable', {
      ...table,
      styles: {
        bodyParagraph: { tabStops: [{ position: 12, type: 'right' }] },
        headerText: { fontWeight: 600 },
        stripedCell: { backgroundColor: '#EFF6FF' },
      },
    }),
    node('custom', { styles: { arbitraryPart: { arbitraryValue: true } } }),
  ])('accepts serializable styles consistently: %j', pluginNode => {
    const schema = { content: [pluginNode] }
    expect(validateResource(schema)).toBe(true)
    expect(validateAi(schema)).toBe(true)
    expect(validateSchema(schema)).toEqual({ errors: [], valid: true })
  })

  it.each([
    [
      'dataTable',
      { ...table, styles: { bodyCell: { height: 20 } } },
      '/styles/bodyCell/height',
    ],
    ['callout', { ...callout, styles: { titel: {} } }, '/styles/titel'],
    [
      'callout',
      { ...callout, styles: { title: { fontSize: true } } },
      '/styles/title/fontSize',
    ],
    [
      'callout',
      { ...callout, styles: { title: { textAlign: 'center' } } },
      '/styles/title/textAlign',
    ],
    ['callout', { ...callout, styles: null }, '/styles'],
    [
      'dataTable',
      { ...table, styles: { headerCell: { fontSize: 12 } } },
      '/styles/headerCell/fontSize',
    ],
    [
      'dataTable',
      { ...table, styles: { bodyText: { bold: 'yes' } } },
      '/styles/bodyText/bold',
    ],
    [
      'dataTable',
      { ...table, styles: { stripedCell: { border: { width: 'broken' } } } },
      '/styles/stripedCell/border/width',
    ],
  ])(
    'rejects invalid styles with useful paths: %s %j',
    (name, options, suffix) => {
      const schema = { content: [node(name, options)] }
      expect(validateResource(schema)).toBe(false)
      expect(validateAi(schema)).toBe(false)
      const result = validateSchema(schema)
      expect(result.valid).toBe(false)
      expect(
        result.errors.some(
          error => error.path === `/content/0/options${suffix}`,
        ),
      ).toBe(true)
    },
  )

  for (const [name, contract] of Object.entries(
    BUILTIN_PLUGIN_OPTION_SCHEMAS,
  )) {
    it(`validates every ${name} part consistently across AI, resource and runtime schemas`, () => {
      const base =
        name === 'callout' ? callout : name === 'dataTable' ? table : {}
      for (const part of Object.keys(contract.properties.styles.properties)) {
        const valid = {
          content: [node(name, { ...base, styles: { [part]: {} } })],
        }
        expect(validateResource(valid)).toBe(true)
        expect(validateAi(valid)).toBe(true)
        expect(validateSchema(valid).valid).toBe(true)
        const invalid = {
          content: [
            node(name, {
              ...base,
              styles: { [part]: { unknownProperty: true } },
            }),
          ],
        }
        expect(validateResource(invalid)).toBe(false)
        expect(validateAi(invalid)).toBe(false)
        expect(validateSchema(invalid).errors).toEqual(
          expect.arrayContaining([
            expect.objectContaining({
              path: `/content/0/options/styles/${part}/unknownProperty`,
            }),
          ]),
        )
      }
      const unknown = {
        content: [node(name, { ...base, styles: { unknownPart: {} } })],
      }
      expect(validateResource(unknown)).toBe(false)
      expect(validateAi(unknown)).toBe(false)
      expect(validateSchema(unknown).valid).toBe(false)
    })
  }

  it.each(['header', 'footer'])(
    'validates styles in all nested %s variants consistently',
    field => {
      for (const variant of ['default', 'first', 'even']) {
        const nested = (pluginNode: unknown) => ({
          content: [
            {
              type: 'sectionBreak',
              config: {
                [field]: {
                  [variant]: { children: ['Literal header', pluginNode] },
                },
              },
            },
          ],
        })
        const valid = nested(
          node('metricCard', {
            items: [],
            styles: { emptyText: { docx: { bold: false } } },
          }),
        )
        expect(validateAi(valid)).toBe(true)
        expect(validateResource(valid)).toBe(true)
        expect(validateSchema(valid).valid).toBe(true)
        const invalid = nested(
          node('metricCard', {
            items: [],
            styles: { emptyText: { fontSize: true } },
          }),
        )
        expect(validateAi(invalid)).toBe(false)
        expect(validateResource(invalid)).toBe(false)
        expect(validateSchema(invalid).errors).toContainEqual(
          expect.objectContaining({
            path: `/content/0/config/${field}/${variant}/children/1/options/styles/emptyText/fontSize`,
          }),
        )
      }
    },
  )

  it('exposes option schemas in help and the AI guide', () => {
    for (const source of createBuiltinPluginSources()) {
      if (source.type !== 'inline') {
        continue
      }
      expect(buildPluginHelp(source.plugin).optionsSchema).toHaveProperty(
        'properties.styles',
      )
    }
    expect(generateSchemaGuide()).toContain(
      'initialization styles → common node style → invocation part styles',
    )
    expect(generateSchemaGuide()).toContain('stripedCell')
  })
  it('renders serializable part styles through the built-in plugin loader', async () => {
    const schema = {
      plugins: createBuiltinPluginSources(),
      content: [
        {
          name: 'callout',
          type: 'plugin' as const,
          options: {
            ...callout,
            styles: { cell: { marginLeft: 0 }, title: { color: '#123ABC' } },
          },
        },
        {
          name: 'dataTable',
          type: 'plugin' as const,
          options: {
            ...table,
            styles: {
              headerText: { color: '#456DEF' },
              stripedCell: { backgroundColor: '#ABCDEF' },
            },
          },
        },
      ],
    }
    const zip = await JSZip.loadAsync(
      await (await renderDocx(schema)).toBuffer(),
    )
    const xml = await zip.file('word/document.xml')?.async('string')
    expect(xml).toContain('<w:color w:val="123ABC"/>')
    expect(xml).toContain('<w:color w:val="456DEF"/>')
    expect(xml).toContain('w:fill="ABCDEF"')
    expect(xml).toContain('<w:left w:type="dxa" w:w="0"/>')
  })
})
