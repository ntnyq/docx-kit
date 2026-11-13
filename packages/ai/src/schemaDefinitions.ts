import {
  CELL_STYLE_SCHEMA,
  PARAGRAPH_STYLE_SCHEMA,
  STYLE_RULE_SCHEMA,
  TEXT_STYLE_SCHEMA,
  UNIT_VALUE_SCHEMA,
} from './styleSchema'

const SHARED_STYLE_SCHEMAS = {
  cellStyle: CELL_STYLE_SCHEMA,
  paragraphStyle: PARAGRAPH_STYLE_SCHEMA,
  styleBorder: STYLE_RULE_SCHEMA.properties.border,
  styleMargin: STYLE_RULE_SCHEMA.properties.margin,
  styleRule: STYLE_RULE_SCHEMA,
  styleTabs: STYLE_RULE_SCHEMA.properties.tabStops,
  textStyle: TEXT_STYLE_SCHEMA,
  unitValue: UNIT_VALUE_SCHEMA,
}

const STYLE_REFERENCES = new Map<object, string>(
  Object.entries(SHARED_STYLE_SCHEMAS).map(([name, schema]) => [schema, name]),
)

/**
 * Shared definitions for self-contained draft-07 document and option schemas.
 * Nested lengths, borders and part styles use references too, so validators
 * compile reusable functions rather than repeatedly inlining large fragments.
 */
export const STYLE_SCHEMA_DEFINITIONS: Record<
  string,
  Record<string, unknown>
> = Object.fromEntries(
  Object.entries(SHARED_STYLE_SCHEMAS).map(([name, schema]) => [
    name,
    referenceStyleSchemas(schema),
  ]),
)

/**
 * Reference canonical style fragments without changing their public metadata.
 * Include STYLE_SCHEMA_DEFINITIONS in the containing document's definitions.
 * The root stays inline so this also compiles a standalone style/option schema.
 */
export function referenceStyleSchemas(
  schema: Record<string, unknown>,
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(schema).map(([key, value]) => [key, referenceValue(value)]),
  )
}

function referenceValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(referenceValue)
  }
  if (value && typeof value === 'object') {
    const name = STYLE_REFERENCES.get(value)
    return name
      ? { $ref: `#/definitions/${name}` }
      : referenceStyleSchemas(value as Record<string, unknown>)
  }
  return value
}
