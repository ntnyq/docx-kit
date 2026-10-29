import Ajv from 'ajv'
import { STYLE_RULE_SCHEMA } from '../schema/styles'
import type { ValidationError } from './validateSchema'

const validate = new Ajv({ allErrors: true }).compile(STYLE_RULE_SCHEMA)

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
