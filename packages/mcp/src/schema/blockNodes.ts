export interface BlockNodeDefinition {
  fields: Record<string, SchemaFieldRule>
  requireOneOf?: readonly string[]
}

export type SchemaFieldKind =
  'array' | 'boolean' | 'number' | 'object' | 'string' | 'unknown'

export interface SchemaFieldRule {
  kind: SchemaFieldKind
  enum?: readonly (number | string)[]
  integer?: boolean
  maximum?: number
  minimum?: number
  nonEmpty?: boolean
  required?: boolean
}

const array = (required = false): SchemaFieldRule => ({
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
    fields: { children: array(true), name: string(true) },
  },
  bulletList: {
    fields: {
      bullet: string(),
      items: array(true),
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
      children: array(true),
      comment: array(true),
      date: string(),
      initials: string(),
    },
  },
  deletedText: {
    fields: {
      author: string(true),
      children: array(true),
      date: string(true),
      revisionId: number({ integer: true, minimum: 0, required: true }),
    },
  },
  footnote: {
    fields: { content: array(true) },
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
      children: array(true),
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
      children: array(true),
      date: string(true),
      revisionId: number({ integer: true, minimum: 0, required: true }),
    },
  },
  math: {
    fields: { children: array(true) },
  },
  numberedList: {
    fields: {
      items: array(true),
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
    fields: { children: array(), text: string() },
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
      children: array(),
      text: string(),
    },
  },
} as const satisfies Record<string, BlockNodeDefinition>

export type BlockNodeType = keyof typeof BLOCK_NODE_DEFINITIONS

export const BLOCK_NODE_TYPES = Object.freeze(
  Object.keys(BLOCK_NODE_DEFINITIONS) as BlockNodeType[],
)

export function buildBlockNodeJsonSchemas(): Record<string, unknown>[] {
  return BLOCK_NODE_TYPES.map(type => {
    const definition = BLOCK_NODE_DEFINITIONS[type]
    const required = ['type']
    const properties: Record<string, unknown> = {
      type: { const: type },
    }

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
  })
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
  return schema
}
