export interface BlockNodeDefinition {
  fields: Record<string, SchemaFieldRule>
  requireOneOf?: readonly string[]
}

export type NodeCollectionKind =
  'block' | 'inline' | 'listItem' | 'math' | 'paragraph' | 'text'

export type SchemaFieldKind =
  'array' | 'boolean' | 'number' | 'object' | 'string' | 'unknown'

export interface SchemaFieldRule {
  kind: SchemaFieldKind
  enum?: readonly (number | string)[]
  integer?: boolean
  items?: NodeCollectionKind
  maximum?: number
  minimum?: number
  nonEmpty?: boolean
  required?: boolean
}

const array = (
  required = false,
  items?: NodeCollectionKind,
): SchemaFieldRule => ({
  items,
  kind: 'array',
  required,
})
const number = (
  options: Omit<SchemaFieldRule, 'kind'> = {},
): SchemaFieldRule => ({ kind: 'number', ...options })
const object = (required = false): SchemaFieldRule => ({
  kind: 'object',
  required,
})
const string = (required = false): SchemaFieldRule => ({
  kind: 'string',
  nonEmpty: required,
  required,
})
const unknown = (required = false): SchemaFieldRule => ({
  kind: 'unknown',
  required,
})

/**
 * Canonical runtime metadata for every public block-node discriminator.
 *
 * MCP validation and the exposed JSON Schema are both generated from this map.
 */
export const BLOCK_NODE_DEFINITIONS = {
  columnBreak: { fields: {} },
  pageBreak: { fields: {} },
  thematicBreak: { fields: {} },
  bookmark: {
    fields: { children: array(true, 'text'), name: string(true) },
  },
  bulletList: {
    fields: {
      bullet: string(),
      items: array(true, 'listItem'),
      level: number({ integer: true, maximum: 8, minimum: 0 }),
    },
  },
  checkbox: {
    fields: {
      alias: string(),
      checked: { kind: 'boolean' },
      checkedState: object(),
      label: string(),
      uncheckedState: object(),
    },
  },
  comment: {
    fields: {
      author: string(true),
      children: array(true, 'inline'),
      comment: array(true, 'paragraph'),
      date: string(),
      initials: string(),
    },
  },
  deletedText: {
    fields: {
      author: string(true),
      children: array(true, 'text'),
      date: string(true),
      revisionId: number({ integer: true, minimum: 0, required: true }),
    },
  },
  footnote: {
    fields: { content: array(true, 'paragraph') },
  },
  heading: {
    fields: {
      text: string(true),
      level: number({
        enum: [1, 2, 3, 4, 5, 6],
        integer: true,
        required: true,
      }),
    },
  },
  hyperlink: {
    requireOneOf: ['anchor', 'url'],
    fields: {
      anchor: string(),
      children: array(true, 'text'),
      url: string(),
    },
  },
  image: {
    fields: {
      alt: string(),
      data: unknown(true),
      floating: unknown(),
      height: unknown(),
      width: unknown(),
      imageType: {
        enum: ['bmp', 'gif', 'jpeg', 'jpg', 'png'],
        kind: 'string',
      },
    },
  },
  insertedText: {
    fields: {
      author: string(true),
      children: array(true, 'text'),
      date: string(true),
      revisionId: number({ integer: true, minimum: 0, required: true }),
    },
  },
  math: {
    fields: { children: array(true, 'math') },
  },
  numberedList: {
    fields: {
      items: array(true, 'listItem'),
      level: number({ integer: true, maximum: 8, minimum: 0 }),
      start: number({ integer: true, minimum: 1 }),
      numberingFormat: {
        kind: 'string',
        enum: [
          'decimal',
          'lowerLetter',
          'lowerRoman',
          'upperLetter',
          'upperRoman',
        ],
      },
    },
  },
  paragraph: {
    fields: { children: array(false, 'inline'), text: string() },
  },
  plugin: {
    fields: { name: string(true), options: unknown(true) },
  },
  sectionBreak: {
    fields: { config: object() },
  },
  table: {
    fields: {
      columns: { kind: 'array', nonEmpty: true, required: true },
      data: array(true),
      header: { kind: 'boolean' },
      striped: { kind: 'boolean' },
    },
  },
  textBox: {
    fields: {
      box: object(true),
      children: array(false, 'inline'),
      text: string(),
    },
  },
} as const satisfies Record<string, BlockNodeDefinition>

export type BlockNodeType = keyof typeof BLOCK_NODE_DEFINITIONS

/**
 * Nested content contracts shared by validation and the exposed JSON Schema.
 */
export const TEXT_NODE_DEFINITION: BlockNodeDefinition = {
  fields: { text: { kind: 'string', required: true } },
}
export const LIST_ITEM_DEFINITION: BlockNodeDefinition = {
  fields: {
    children: array(false, 'inline'),
    level: number({ integer: true, maximum: 8, minimum: 0 }),
    text: string(),
  },
}
export const INLINE_NODE_TYPES = [
  'bookmark',
  'checkbox',
  'comment',
  'deletedText',
  'footnote',
  'hyperlink',
  'image',
  'insertedText',
  'math',
] as const satisfies readonly BlockNodeType[]
export const MATH_NODE_DEFINITIONS: Record<string, BlockNodeDefinition> = {
  text: TEXT_NODE_DEFINITION,
  fraction: {
    fields: {
      denominator: array(true, 'math'),
      numerator: array(true, 'math'),
    },
  },
  function: {
    fields: { arguments: array(true, 'math'), name: array(true, 'math') },
  },
  integral: {
    fields: {
      children: array(true, 'math'),
      subScript: array(false, 'math'),
      superScript: array(false, 'math'),
    },
  },
  radical: {
    fields: { children: array(true, 'math'), degree: array(false, 'math') },
  },
  script: {
    fields: {
      children: array(true, 'math'),
      subScript: array(false, 'math'),
      superScript: array(false, 'math'),
    },
  },
  sum: {
    fields: {
      children: array(true, 'math'),
      subScript: array(false, 'math'),
      superScript: array(false, 'math'),
    },
  },
}

export const BLOCK_NODE_TYPES = Object.freeze(
  Object.keys(BLOCK_NODE_DEFINITIONS) as BlockNodeType[],
)

export function buildBlockNodeJsonSchemas(): Record<string, unknown>[] {
  return BLOCK_NODE_TYPES.map(type =>
    buildNodeJsonSchema(BLOCK_NODE_DEFINITIONS[type], type),
  )
}

export function buildNestedNodeJsonSchemas(): Record<string, unknown> {
  const text = buildNodeJsonSchema(TEXT_NODE_DEFINITION, 'text')
  return {
    textNode: { anyOf: [{ type: 'string' }, text] },
    inlineNode: {
      oneOf: [
        text,
        ...INLINE_NODE_TYPES.map(type =>
          buildNodeJsonSchema(BLOCK_NODE_DEFINITIONS[type], type),
        ),
      ],
    },
    listItemNode: {
      anyOf: [{ type: 'string' }, buildNodeJsonSchema(LIST_ITEM_DEFINITION)],
    },
    mathNode: {
      oneOf: Object.entries(MATH_NODE_DEFINITIONS).map(([type, definition]) =>
        buildNodeJsonSchema(definition, type),
      ),
    },
    paragraphNode: {
      anyOf: [
        { type: 'string' },
        buildNodeJsonSchema(BLOCK_NODE_DEFINITIONS.paragraph, 'paragraph'),
      ],
    },
  }
}

function buildNodeJsonSchema(
  definition: BlockNodeDefinition,
  type?: string,
): Record<string, unknown> {
  const required = type ? ['type'] : []
  const properties: Record<string, unknown> = type
    ? { type: { const: type } }
    : {}

  for (const [field, rule] of Object.entries(definition.fields)) {
    properties[field] = fieldRuleToJsonSchema(rule)
    if (rule.required) {
      required.push(field)
    }
  }

  const schema: Record<string, unknown> = {
    additionalProperties: true,
    properties,
    required,
    type: 'object',
  }

  if ('requireOneOf' in definition && definition.requireOneOf) {
    schema.anyOf = definition.requireOneOf.map(field => ({
      required: [field],
    }))
  }

  return schema
}

function fieldRuleToJsonSchema(rule: SchemaFieldRule): Record<string, unknown> {
  if (rule.kind === 'unknown') {
    return {}
  }

  const schema: Record<string, unknown> = { type: rule.kind }
  if (rule.enum) {
    schema.enum = rule.enum
  }
  if (rule.integer) {
    schema.type = 'integer'
  }
  if (rule.maximum !== undefined) {
    schema.maximum = rule.maximum
  }
  if (rule.minimum !== undefined) {
    schema.minimum = rule.minimum
  }
  if (rule.nonEmpty && rule.kind === 'array') {
    schema.minItems = 1
  }
  if (rule.nonEmpty && rule.kind === 'string') {
    schema.minLength = 1
  }
  if (rule.items) {
    schema.items = { $ref: `#/definitions/${rule.items}Node` }
  }
  return schema
}
