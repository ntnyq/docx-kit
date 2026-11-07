import {
  CELL_STYLE_SCHEMA,
  PARAGRAPH_STYLE_SCHEMA,
  TEXT_STYLE_SCHEMA,
} from './styleSchema'

const string = { type: 'string' }
const boolean = { type: 'boolean' }
const styleDescription =
  'Per-part overrides. Precedence: built-in/theme defaults, plugin initialization styles, common node style, then these part styles. Undefined fields retain defaults; false and zero apply. Theme tokens are supported.'

/**
 * Add a strict part-style contract without restricting other plugin options.
 */
function pluginOptions(
  parts: Record<string, unknown>,
  properties: Record<string, unknown> = {},
  required?: string[],
  conditions?: object[],
) {
  return {
    additionalProperties: true,
    ...(conditions ? { allOf: conditions } : {}),
    ...(required ? { required } : {}),
    type: 'object',
    properties: {
      ...properties,
      styles: {
        additionalProperties: false,
        description: styleDescription,
        properties: parts,
        type: 'object',
      },
    },
  }
}

/**
 * Serializable option contracts shared by AI tools, MCP help and validation.
 * Custom plugins remain extensible.
 */
export const BUILTIN_PLUGIN_OPTION_SCHEMAS: Record<
  string,
  ReturnType<typeof pluginOptions>
> = {
  badge: pluginOptions({
    paragraph: PARAGRAPH_STYLE_SCHEMA,
    text: TEXT_STYLE_SCHEMA,
  }),
  barcode: pluginOptions({
    caption: TEXT_STYLE_SCHEMA,
    captionParagraph: PARAGRAPH_STYLE_SCHEMA,
    imageParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  bibliography: pluginOptions(
    {
      citation: TEXT_STYLE_SCHEMA,
      citationParagraph: PARAGRAPH_STYLE_SCHEMA,
      entry: TEXT_STYLE_SCHEMA,
      entryParagraph: PARAGRAPH_STYLE_SCHEMA,
      title: TEXT_STYLE_SCHEMA,
      titleParagraph: PARAGRAPH_STYLE_SCHEMA,
    },
    {
      bookmarkPrefix: string,
      hyperlink: boolean,
      ids: { items: { minLength: 1, type: 'string' }, type: 'array' },
      mode: { enum: ['citation', 'list'] },
      prefix: string,
      suffix: string,
      title: string,
      entries: {
        type: 'array',
        items: {
          required: ['id', 'title'],
          type: 'object',
          properties: {
            authors: { items: string, type: 'array' },
            doi: string,
            id: { minLength: 1, type: 'string' },
            publisher: string,
            title: { pattern: '\\S', type: 'string' },
            url: string,
            year: string,
          },
        },
      },
    },
    undefined,
    [
      {
        if: { properties: { mode: { const: 'citation' } }, required: ['mode'] },
        then: { required: ['ids'] },
      },
    ],
  ),
  callout: pluginOptions(
    {
      cell: CELL_STYLE_SCHEMA,
      content: TEXT_STYLE_SCHEMA,
      paragraph: PARAGRAPH_STYLE_SCHEMA,
      title: TEXT_STYLE_SCHEMA,
    },
    {
      content: string,
      title: string,
      type: { enum: ['danger', 'info', 'success', 'warning'] },
    },
    ['content', 'type'],
  ),
  changelog: pluginOptions({
    bodyCell: CELL_STYLE_SCHEMA,
    bodyParagraph: PARAGRAPH_STYLE_SCHEMA,
    bodyText: TEXT_STYLE_SCHEMA,
    emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyText: TEXT_STYLE_SCHEMA,
    headerCell: CELL_STYLE_SCHEMA,
    headerParagraph: PARAGRAPH_STYLE_SCHEMA,
    headerText: TEXT_STYLE_SCHEMA,
    title: TEXT_STYLE_SCHEMA,
    titleParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  checklist: pluginOptions({
    emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyText: TEXT_STYLE_SCHEMA,
    group: TEXT_STYLE_SCHEMA,
    groupParagraph: PARAGRAPH_STYLE_SCHEMA,
    item: TEXT_STYLE_SCHEMA,
    itemParagraph: PARAGRAPH_STYLE_SCHEMA,
    note: TEXT_STYLE_SCHEMA,
    noteParagraph: PARAGRAPH_STYLE_SCHEMA,
    owner: TEXT_STYLE_SCHEMA,
    title: TEXT_STYLE_SCHEMA,
    titleParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  codeBlock: pluginOptions({
    code: TEXT_STYLE_SCHEMA,
    lineNumber: TEXT_STYLE_SCHEMA,
    paragraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  coverPage: pluginOptions({
    author: TEXT_STYLE_SCHEMA,
    authorParagraph: PARAGRAPH_STYLE_SCHEMA,
    date: TEXT_STYLE_SCHEMA,
    dateParagraph: PARAGRAPH_STYLE_SCHEMA,
    organization: TEXT_STYLE_SCHEMA,
    organizationParagraph: PARAGRAPH_STYLE_SCHEMA,
    ruleParagraph: PARAGRAPH_STYLE_SCHEMA,
    subtitle: TEXT_STYLE_SCHEMA,
    subtitleParagraph: PARAGRAPH_STYLE_SCHEMA,
    title: TEXT_STYLE_SCHEMA,
    titleParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  crossReference: pluginOptions(
    { paragraph: PARAGRAPH_STYLE_SCHEMA, text: TEXT_STYLE_SCHEMA },
    {
      fallback: string,
      hyperlink: boolean,
      kind: { enum: ['number', 'page', 'text'] },
      prefix: string,
      suffix: string,
      target: {
        pattern: '^[A-Za-z][A-Za-z0-9_]{0,39}(?![\\s\\S])',
        type: 'string',
      },
    },
    ['target'],
  ),
  dataTable: pluginOptions(
    {
      bodyCell: CELL_STYLE_SCHEMA,
      bodyParagraph: PARAGRAPH_STYLE_SCHEMA,
      bodyText: TEXT_STYLE_SCHEMA,
      emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
      emptyText: TEXT_STYLE_SCHEMA,
      headerCell: CELL_STYLE_SCHEMA,
      headerParagraph: PARAGRAPH_STYLE_SCHEMA,
      headerText: TEXT_STYLE_SCHEMA,
      stripedCell: CELL_STYLE_SCHEMA,
    },
    {
      bordered: boolean,
      labels: { additionalProperties: string, type: 'object' },
      striped: boolean,
      align: {
        additionalProperties: { enum: ['center', 'left', 'right'] },
        type: 'object',
      },
      data: {
        items: { additionalProperties: true, type: 'object' },
        type: 'array',
      },
      format: {
        type: 'object',
        additionalProperties: {
          enum: ['currency', 'date', 'number', 'percent'],
        },
      },
    },
    ['data'],
  ),
  divider: pluginOptions({
    paragraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  echarts: pluginOptions({
    caption: TEXT_STYLE_SCHEMA,
    captionParagraph: PARAGRAPH_STYLE_SCHEMA,
    imageParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  figure: pluginOptions({
    caption: TEXT_STYLE_SCHEMA,
    captionParagraph: PARAGRAPH_STYLE_SCHEMA,
    imageParagraph: PARAGRAPH_STYLE_SCHEMA,
    source: TEXT_STYLE_SCHEMA,
    sourceParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  glossary: pluginOptions(
    {
      bodyCell: CELL_STYLE_SCHEMA,
      bodyParagraph: PARAGRAPH_STYLE_SCHEMA,
      definition: TEXT_STYLE_SCHEMA,
      fullForm: TEXT_STYLE_SCHEMA,
      group: TEXT_STYLE_SCHEMA,
      groupParagraph: PARAGRAPH_STYLE_SCHEMA,
      headerCell: CELL_STYLE_SCHEMA,
      headerParagraph: PARAGRAPH_STYLE_SCHEMA,
      headerText: TEXT_STYLE_SCHEMA,
      term: TEXT_STYLE_SCHEMA,
      title: TEXT_STYLE_SCHEMA,
      titleParagraph: PARAGRAPH_STYLE_SCHEMA,
    },
    {
      groupBy: { enum: ['group', 'none'] },
      locale: string,
      sort: { enum: ['none', 'term'] },
      title: string,
      entries: {
        type: 'array',
        items: {
          required: ['term', 'definition'],
          type: 'object',
          properties: {
            definition: { pattern: '\\S', type: 'string' },
            fullForm: string,
            group: string,
            term: { pattern: '\\S', type: 'string' },
          },
        },
      },
      labels: {
        additionalProperties: false,
        properties: { definition: string, fullForm: string, term: string },
        type: 'object',
      },
    },
    ['entries'],
  ),
  invoice: pluginOptions({
    bodyCell: CELL_STYLE_SCHEMA,
    bodyParagraph: PARAGRAPH_STYLE_SCHEMA,
    bodyText: TEXT_STYLE_SCHEMA,
    headerCell: CELL_STYLE_SCHEMA,
    headerParagraph: PARAGRAPH_STYLE_SCHEMA,
    headerText: TEXT_STYLE_SCHEMA,
    metaText: TEXT_STYLE_SCHEMA,
    notes: TEXT_STYLE_SCHEMA,
    notesLabel: TEXT_STYLE_SCHEMA,
    notesParagraph: PARAGRAPH_STYLE_SCHEMA,
    partyLabel: TEXT_STYLE_SCHEMA,
    partyParagraph: PARAGRAPH_STYLE_SCHEMA,
    partyText: TEXT_STYLE_SCHEMA,
    subtotal: TEXT_STYLE_SCHEMA,
    tax: TEXT_STYLE_SCHEMA,
    title: TEXT_STYLE_SCHEMA,
    titleParagraph: PARAGRAPH_STYLE_SCHEMA,
    total: TEXT_STYLE_SCHEMA,
    totalsParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  letterhead: pluginOptions({
    address: TEXT_STYLE_SCHEMA,
    addressParagraph: PARAGRAPH_STYLE_SCHEMA,
    companyName: TEXT_STYLE_SCHEMA,
    companyParagraph: PARAGRAPH_STYLE_SCHEMA,
    contact: TEXT_STYLE_SCHEMA,
    contactParagraph: PARAGRAPH_STYLE_SCHEMA,
    ruleParagraph: PARAGRAPH_STYLE_SCHEMA,
    tagline: TEXT_STYLE_SCHEMA,
    taglineParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  meetingMinutes: pluginOptions({
    bodyCell: CELL_STYLE_SCHEMA,
    bodyParagraph: PARAGRAPH_STYLE_SCHEMA,
    bodyText: TEXT_STYLE_SCHEMA,
    headerCell: CELL_STYLE_SCHEMA,
    headerParagraph: PARAGRAPH_STYLE_SCHEMA,
    headerText: TEXT_STYLE_SCHEMA,
    metaLabel: TEXT_STYLE_SCHEMA,
    metaParagraph: PARAGRAPH_STYLE_SCHEMA,
    metaText: TEXT_STYLE_SCHEMA,
    title: TEXT_STYLE_SCHEMA,
    titleParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  metricCard: pluginOptions({
    cell: CELL_STYLE_SCHEMA,
    change: TEXT_STYLE_SCHEMA,
    changeParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyText: TEXT_STYLE_SCHEMA,
    label: TEXT_STYLE_SCHEMA,
    labelParagraph: PARAGRAPH_STYLE_SCHEMA,
    paddingParagraph: PARAGRAPH_STYLE_SCHEMA,
    unit: TEXT_STYLE_SCHEMA,
    value: TEXT_STYLE_SCHEMA,
    valueParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  pageNumber: pluginOptions({
    paragraph: PARAGRAPH_STYLE_SCHEMA,
    text: TEXT_STYLE_SCHEMA,
  }),
  propertyTable: pluginOptions({
    emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyText: TEXT_STYLE_SCHEMA,
    keyCell: CELL_STYLE_SCHEMA,
    keyParagraph: PARAGRAPH_STYLE_SCHEMA,
    keyText: TEXT_STYLE_SCHEMA,
    stripedCell: CELL_STYLE_SCHEMA,
    valueCell: CELL_STYLE_SCHEMA,
    valueParagraph: PARAGRAPH_STYLE_SCHEMA,
    valueText: TEXT_STYLE_SCHEMA,
  }),
  qrcode: pluginOptions({
    caption: TEXT_STYLE_SCHEMA,
    captionParagraph: PARAGRAPH_STYLE_SCHEMA,
    imageParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  signatureBlock: pluginOptions({
    cell: CELL_STYLE_SCHEMA,
    date: TEXT_STYLE_SCHEMA,
    dateParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyText: TEXT_STYLE_SCHEMA,
    label: TEXT_STYLE_SCHEMA,
    labelParagraph: PARAGRAPH_STYLE_SCHEMA,
    paddingParagraph: PARAGRAPH_STYLE_SCHEMA,
    signature: TEXT_STYLE_SCHEMA,
    signatureParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  timeline: pluginOptions({
    connector: TEXT_STYLE_SCHEMA,
    connectorCell: CELL_STYLE_SCHEMA,
    connectorParagraph: PARAGRAPH_STYLE_SCHEMA,
    contentCell: CELL_STYLE_SCHEMA,
    contentParagraph: PARAGRAPH_STYLE_SCHEMA,
    date: TEXT_STYLE_SCHEMA,
    dateCell: CELL_STYLE_SCHEMA,
    dateParagraph: PARAGRAPH_STYLE_SCHEMA,
    description: TEXT_STYLE_SCHEMA,
    emptyParagraph: PARAGRAPH_STYLE_SCHEMA,
    emptyText: TEXT_STYLE_SCHEMA,
    title: TEXT_STYLE_SCHEMA,
  }),
  toc: pluginOptions({
    title: TEXT_STYLE_SCHEMA,
    titleParagraph: PARAGRAPH_STYLE_SCHEMA,
  }),
  watermark: pluginOptions({
    paragraph: PARAGRAPH_STYLE_SCHEMA,
    text: TEXT_STYLE_SCHEMA,
  }),
}

/**
 * Constrain known plugin options while allowing arbitrary other nodes/plugins.
 */
export const BUILTIN_PLUGIN_SCHEMA_CONDITIONS = Object.entries(
  BUILTIN_PLUGIN_OPTION_SCHEMAS,
).map(([name, schema]) => ({
  if: {
    properties: { name: { const: name }, type: { const: 'plugin' } },
    required: ['type', 'name'],
  },
  then: {
    properties: { options: schema },
    required: ['options'],
  },
}))

const headerFooterSchema = {
  additionalProperties: true,
  type: 'object',
  properties: Object.fromEntries(
    ['default', 'even', 'first'].map(variant => [
      variant,
      {
        additionalProperties: true,
        required: ['children'],
        type: 'object',
        properties: {
          children: {
            type: 'array',
            items: {
              anyOf: [{ type: 'string' }, { $ref: '#/definitions/blockNode' }],
            },
          },
        },
      },
    ]),
  ),
}

/**
 * Reuse the document's block-node contract for plugins nested in section
 * headers and footers. Consumers must define `definitions.blockNode`.
 */
export const SECTION_BREAK_SCHEMA_CONDITION = {
  if: { properties: { type: { const: 'sectionBreak' } }, required: ['type'] },
  then: {
    properties: {
      config: {
        additionalProperties: true,
        properties: { footer: headerFooterSchema, header: headerFooterSchema },
        type: 'object',
      },
    },
  },
}
