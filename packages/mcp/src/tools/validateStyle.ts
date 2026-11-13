import Ajv from 'ajv'
import {
  referenceStyleSchemas,
  STYLE_RULE_SCHEMA,
  STYLE_SCHEMA_DEFINITIONS,
} from '../schema/styles'
import type { ValidationError } from './validateSchema'

const validate = new Ajv({ allErrors: true }).compile({
  ...referenceStyleSchemas(STYLE_RULE_SCHEMA),
  definitions: STYLE_SCHEMA_DEFINITIONS,
})

/**
 * Validate the same style schema exposed to MCP clients without echoing values.
 */
export function validateStyle(
  value: unknown,
  path: string,
  errors: ValidationError[],
) {
  if (!validate(value)) {
    for (const error of validate.errors ?? []) {
      errors.push({
        message: `Style ${error.message ?? 'is invalid'}`,
        path: `${path}${error.instancePath}`,
      })
    }
  }
}
