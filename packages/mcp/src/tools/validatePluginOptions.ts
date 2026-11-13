import {
  BUILTIN_PLUGIN_OPTION_SCHEMAS,
  referenceStyleSchemas,
  STYLE_SCHEMA_DEFINITIONS,
} from '@docxkit/ai'
import Ajv from 'ajv'
import type { ValidationError } from './validateSchema'

const ajv = new Ajv({ allErrors: true })
const validators = new Map(
  Object.entries(BUILTIN_PLUGIN_OPTION_SCHEMAS).map(([name, schema]) => [
    name,
    ajv.compile({
      ...referenceStyleSchemas(schema),
      definitions: STYLE_SCHEMA_DEFINITIONS,
    }),
  ]),
)

/**
 * Validate known built-in options using the contract exposed in AI/MCP help.
 * Custom plugin options retain their existing extensible behavior.
 */
export function validatePluginOptions(
  name: unknown,
  options: unknown,
  path: string,
  errors: ValidationError[],
) {
  const validate = typeof name === 'string' ? validators.get(name) : undefined
  if (!validate || validate(options)) {
    return
  }
  for (const error of validate.errors ?? []) {
    const field =
      error.keyword === 'additionalProperties'
        ? error.params.additionalProperty
        : error.keyword === 'required'
          ? error.params.missingProperty
          : undefined
    const suffix =
      typeof field === 'string'
        ? `/${field.replaceAll('~', '~0').replaceAll('/', '~1')}`
        : ''
    errors.push({
      message: `Plugin options ${error.message ?? 'are invalid'}`,
      path: `${path}${error.instancePath}${suffix}`,
    })
  }
}
