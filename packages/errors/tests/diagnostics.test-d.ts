import { expectTypeOf } from 'vitest'
import { createDocxKitError, DIAGNOSTIC_CODES, ERROR_CODES } from '../src/index'
import type { ErrorParams } from '../src/index'

expectTypeOf(
  ERROR_CODES.PLUGIN_NOT_REGISTERED,
).toEqualTypeOf<'PLUGIN_NOT_REGISTERED'>()
expectTypeOf(
  DIAGNOSTIC_CODES.PLUGIN_NOT_REGISTERED,
).toEqualTypeOf<'DOCXKIT_PLUGIN_NOT_REGISTERED'>()
expectTypeOf<ErrorParams<'PLUGIN_NOT_REGISTERED'>>().toExtend<{
  pluginName: string
}>()

// @ts-expect-error Plugin names are required by this diagnostic.
createDocxKitError('PLUGIN_NOT_REGISTERED', {})
// @ts-expect-error Parameter types depend on the selected code.
createDocxKitError('EXPORT_FAILED', { pluginName: 'qrcode' })
// @ts-expect-error Unknown codes still use the legacy constructor.
createDocxKitError('CUSTOM', {})
// @ts-expect-error A plugin loading repair requires an explicit reason.
createDocxKitError('PLUGIN_LOAD_FAILED', { detail: 'Failed' })
// @ts-expect-error Only supported export formats are accepted.
createDocxKitError('EXPORT_FAILED', { format: 'pdf' })
