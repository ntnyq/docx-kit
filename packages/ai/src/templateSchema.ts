import type { AiTemplateSchema } from './types'

/**
 * A template input error, located by a JSON Pointer without echoing its value.
 */
export interface TemplateValidationError {
  message: string
  path: string
}

/**
 * Convert template field definitions to standard JSON Schema, moving required
 * markers onto their containing objects, including objects nested in arrays.
 */
export function toTemplateJsonSchema(
  schema: AiTemplateSchema,
): Record<string, unknown> {
  const { required: _, ...definition } = schema
  if (schema.type === 'array') {
    return {
      ...definition,
      items: toTemplateJsonSchema(schema.items),
    }
  }
  if (schema.type === 'object') {
    return {
      ...definition,
      properties: Object.fromEntries(
        Object.entries(schema.properties).map(([key, field]) => [
          key,
          toTemplateJsonSchema(field),
        ]),
      ),
      required: Object.entries(schema.properties)
        .filter(([, field]) => field.required)
        .map(([key]) => key),
    }
  }
  return definition
}

/**
 * Validate unknown template input before calling its typed generator.
 * Extra properties are allowed; declared properties and array items must match.
 */
export function validateTemplateData(
  schema: AiTemplateSchema,
  value: unknown,
): TemplateValidationError[] {
  const errors: TemplateValidationError[] = []
  validateValue(schema, value, '', errors)
  return errors
}

function validateValue(
  schema: AiTemplateSchema,
  value: unknown,
  path: string,
  errors: TemplateValidationError[],
) {
  switch (schema.type) {
    case 'array':
      if (Array.isArray(value)) {
        value.forEach((item, index) =>
          validateValue(schema.items, item, `${path}/${index}`, errors),
        )
        return
      }
      break
    case 'boolean':
      if (typeof value === 'boolean') {
        return
      }
      break
    case 'number':
      if (typeof value === 'number' && Number.isFinite(value)) {
        return
      }
      break
    case 'object':
      if (
        typeof value === 'object'
        && value !== null
        && !Array.isArray(value)
      ) {
        for (const [key, field] of Object.entries(schema.properties)) {
          const fieldPath = `${path}/${key.replaceAll('~', '~0').replaceAll('/', '~1')}`
          if (Object.hasOwn(value, key)) {
            validateValue(field, Reflect.get(value, key), fieldPath, errors)
          } else if (field.required) {
            errors.push({
              message: 'Required field is missing',
              path: fieldPath,
            })
          }
        }
        return
      }
      break
    case 'string':
      if (typeof value === 'string') {
        if (schema.enum && !schema.enum.includes(value)) {
          errors.push({
            message: `Expected one of: ${schema.enum.join(', ')}`,
            path,
          })
        }
        return
      }
      break
  }
  errors.push({ message: `Expected ${schema.type}`, path })
}
