/**
 * MCP tool: validate a docx-kit JSON schema.
 *
 * @module mcp-server/tools/validateSchema
 */

import {
  BLOCK_NODE_DEFINITIONS,
  INLINE_NODE_TYPES,
  LIST_ITEM_DEFINITION,
  MATH_NODE_DEFINITIONS,
  TEXT_NODE_DEFINITION,
} from '../schema/blockNodes'
import type {
  BlockNodeDefinition,
  BlockNodeType,
  NodeCollectionKind,
  SchemaFieldRule,
} from '../schema/blockNodes'

/**
 * Output from the validate_schema MCP tool.
 */
export interface ValidateSchemaOutput {
  /**
   * List of validation errors (empty if valid).
   */
  errors: ValidationError[]
  /**
   * Whether the schema is valid.
   */
  valid: boolean
}

/**
 * Validation error detail.
 */
export interface ValidationError {
  /**
   * Error description.
   */
  message: string
  /**
   * JSON path to the invalid field.
   */
  path: string
}

/**
 * MCP tool definition for `validate_schema`.
 *
 * Validates a docx-kit JSON schema and returns
 * detailed error information.
 */
export const validateSchemaToolDefinition = {
  name: 'validate_schema',
  description:
    'Validate a docx-kit JSON schema. Checks node types, required fields, and structural correctness.',
  inputSchema: {
    required: ['schema'],
    type: 'object',
    properties: {
      schema: {
        description: 'The DocxSchema JSON object to validate',
        type: 'object',
      },
    },
  },
}

/**
 * Validate a DocxSchema object.
 *
 * Checks for:
 * - Required `content` array
 * - Valid node types in content
 * - Required fields per node type
 *
 * @param schema - — The schema to validate
 * @returns Validation result with errors
 */
export function validateSchema(schema: unknown): ValidateSchemaOutput {
  const errors: ValidationError[] = []
  const doc = schema as Record<string, unknown> | undefined
  const content = doc?.content

  if (!content) {
    errors.push({
      message: 'Required field "content" is missing',
      path: '/content',
    })
  } else if (Array.isArray(content)) {
    for (const [i, element] of content.entries()) {
      const node = element as Record<string, unknown>
      if (typeof node !== 'object' || node === null) {
        errors.push({
          message: `Node at index ${i} must be an object`,
          path: `/content/${i}`,
        })
        continue
      }

      const nodeType = node.type
      if (typeof nodeType !== 'string' || nodeType.length === 0) {
        errors.push({
          message: `Node at index ${i} missing required "type" field`,
          path: `/content/${i}/type`,
        })
        continue
      }

      if (!Object.hasOwn(BLOCK_NODE_DEFINITIONS, nodeType)) {
        errors.push({
          message: `Invalid node type "${nodeType}" at index ${i}`,
          path: `/content/${i}/type`,
        })
        continue
      }

      validateNode(node, nodeType as BlockNodeType, `/content/${i}`, errors)
    }
  } else {
    errors.push({
      message: 'Field "content" must be an array',
      path: '/content',
    })
  }

  return {
    errors,
    valid: errors.length === 0,
  }
}

function article(kind: string): string {
  return kind === 'array' || kind === 'object' ? 'an' : 'a'
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function matchesKind(value: unknown, kind: SchemaFieldRule['kind']): boolean {
  switch (kind) {
    case 'array':
      return Array.isArray(value)
    case 'boolean':
      return typeof value === 'boolean'
    case 'number':
      return typeof value === 'number' && Number.isFinite(value)
    case 'object':
      return (
        Boolean(value) && typeof value === 'object' && !Array.isArray(value)
      )
    case 'string':
      return typeof value === 'string'
    case 'unknown':
      return true
  }
}

function validateCollection(
  nodes: unknown[],
  kind: NodeCollectionKind,
  path: string,
  errors: ValidationError[],
  depth: number,
): void {
  if (depth > 64) {
    errors.push({ message: 'Content nesting exceeds 64 levels', path })
    return
  }
  for (const [index, node] of nodes.entries()) {
    const nodePath = `${path}/${index}`
    if (
      typeof node === 'string'
      && ['listItem', 'paragraph', 'text'].includes(kind)
    ) {
      continue
    }
    if (!isRecord(node)) {
      errors.push({ message: `Expected a ${kind} node`, path: nodePath })
      continue
    }
    if (kind === 'listItem') {
      validateDefinition(node, LIST_ITEM_DEFINITION, nodePath, errors, depth)
      continue
    }
    const type = node.type
    if (typeof type === 'string') {
      if (kind === 'math' && Object.hasOwn(MATH_NODE_DEFINITIONS, type)) {
        validateDefinition(
          node,
          MATH_NODE_DEFINITIONS[type],
          nodePath,
          errors,
          depth,
        )
        continue
      }
      if (type === 'text' && (kind === 'text' || kind === 'inline')) {
        validateDefinition(node, TEXT_NODE_DEFINITION, nodePath, errors, depth)
        continue
      }
      const isAllowed =
        kind === 'block'
        || (kind === 'paragraph' && type === 'paragraph')
        || (kind === 'inline'
          && (INLINE_NODE_TYPES as readonly string[]).includes(type))
      if (isAllowed && Object.hasOwn(BLOCK_NODE_DEFINITIONS, type)) {
        validateNode(node, type as BlockNodeType, nodePath, errors, depth)
        continue
      }
    }
    errors.push({
      message: `Invalid ${kind} node type: ${String(type)}`,
      path: `${nodePath}/type`,
    })
  }
}

function validateDefinition(
  node: Record<string, unknown>,
  definition: BlockNodeDefinition,
  path: string,
  errors: ValidationError[],
  depth: number,
): void {
  for (const [field, rule] of Object.entries(definition.fields)) {
    const hasField = Object.hasOwn(node, field)
    if (rule.required && !hasField) {
      errors.push({
        message: `${String(node.type ?? 'List item')} node missing required "${field}" field`,
        path: `${path}/${field}`,
      })
      continue
    }
    if (hasField) {
      validateField(node[field], field, rule, path, errors, depth)
    }
  }

  if (
    'requireOneOf' in definition
    && definition.requireOneOf
    && !definition.requireOneOf.some(field => isNonEmptyString(node[field]))
  ) {
    errors.push({
      message: `${String(node.type)} node requires one of: ${definition.requireOneOf.join(', ')}`,
      path,
    })
  }
}

function validateField(
  value: unknown,
  field: string,
  rule: SchemaFieldRule,
  path: string,
  errors: ValidationError[],
  depth = 0,
): void {
  if (rule.kind !== 'unknown' && !matchesKind(value, rule.kind)) {
    errors.push({
      message: `Field "${field}" must be ${article(rule.kind)} ${rule.kind}`,
      path: `${path}/${field}`,
    })
    return
  }

  if (rule.enum && !rule.enum.includes(value as never)) {
    errors.push({
      message: `Field "${field}" must be one of: ${rule.enum.join(', ')}`,
      path: `${path}/${field}`,
    })
  }

  if (typeof value === 'number') {
    if (rule.integer && !Number.isInteger(value)) {
      errors.push({
        message: `Field "${field}" must be an integer`,
        path: `${path}/${field}`,
      })
    }
    if (rule.minimum !== undefined && value < rule.minimum) {
      errors.push({
        message: `Field "${field}" must be at least ${rule.minimum}`,
        path: `${path}/${field}`,
      })
    }
    if (rule.maximum !== undefined && value > rule.maximum) {
      errors.push({
        message: `Field "${field}" must be at most ${rule.maximum}`,
        path: `${path}/${field}`,
      })
    }
  }

  if (rule.nonEmpty && Array.isArray(value) && value.length === 0) {
    errors.push({
      message: `Field "${field}" must not be empty`,
      path: `${path}/${field}`,
    })
  }
  if (rule.nonEmpty && typeof value === 'string' && value.length === 0) {
    errors.push({
      message: `Field "${field}" must not be empty`,
      path: `${path}/${field}`,
    })
  }
  if (rule.items && Array.isArray(value)) {
    validateCollection(value, rule.items, `${path}/${field}`, errors, depth + 1)
  }
}

function validateHeaderFooters(
  config: unknown,
  path: string,
  errors: ValidationError[],
  depth: number,
): void {
  if (!isRecord(config)) {
    return
  }
  for (const field of ['header', 'footer']) {
    const variants = config[field]
    if (variants === undefined) {
      continue
    }
    if (!isRecord(variants)) {
      errors.push({
        message: `${field} must be an object`,
        path: `${path}/${field}`,
      })
      continue
    }
    for (const key of ['default', 'first', 'even']) {
      const content = variants[key]
      if (content === undefined) {
        continue
      }
      const contentPath = `${path}/${field}/${key}/children`
      if (!isRecord(content) || !Array.isArray(content.children)) {
        errors.push({
          message: 'Header/footer children must be an array',
          path: contentPath,
        })
        continue
      }
      // Header/footer strings are shorthand paragraphs, unlike ordinary blocks.
      validateCollection(
        content.children.map(child =>
          typeof child === 'string'
            ? { text: child, type: 'paragraph' }
            : child,
        ),
        'block',
        contentPath,
        errors,
        depth,
      )
    }
  }
}

function validateNode(
  node: Record<string, unknown>,
  nodeType: BlockNodeType,
  path: string,
  errors: ValidationError[],
  depth = 0,
): void {
  const definition = BLOCK_NODE_DEFINITIONS[nodeType]

  validateDefinition(node, definition, path, errors, depth)

  if (nodeType === 'sectionBreak') {
    validateHeaderFooters(node.config, `${path}/config`, errors, depth + 1)
  }
  if (nodeType === 'table') {
    validateTableColumns(node.columns, path, errors)
  }
  if (nodeType === 'textBox') {
    const box = node.box as Record<string, unknown> | undefined
    if (box && !Object.hasOwn(box, 'width')) {
      errors.push({
        message: 'textBox box missing required "width" field',
        path: `${path}/box/width`,
      })
    }
  }
}

function validateTableColumns(
  value: unknown,
  path: string,
  errors: ValidationError[],
): void {
  if (!Array.isArray(value)) {
    return
  }
  for (const [index, column] of value.entries()) {
    if (!column || typeof column !== 'object' || Array.isArray(column)) {
      errors.push({
        message: 'Table column must be an object',
        path: `${path}/columns/${index}`,
      })
      continue
    }
    const record = column as Record<string, unknown>
    for (const field of ['key', 'title']) {
      if (!isNonEmptyString(record[field])) {
        errors.push({
          message: `Table column missing required "${field}" string`,
          path: `${path}/columns/${index}/${field}`,
        })
      }
    }
  }
}
