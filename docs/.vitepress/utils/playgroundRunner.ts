import * as docx from 'docx'
import * as docxKit from 'docx-kit'
import { prepareCode } from './prepareCode'
import type { DocxKitConfig } from 'docx-kit'

export interface PlaygroundRuntimeOverrides {
  createDocx?: (config?: DocxKitConfig) => unknown
}

/**
 * Execute playground source against the complete browser runtime exports.
 *
 * Keeping `docx-kit` as a namespace makes runtime availability match the
 * declarations exposed to Monaco, including every built-in plugin.
 */
export async function executePlaygroundCode(
  raw: string,
  overrides: PlaygroundRuntimeOverrides = {},
): Promise<unknown> {
  const source = prepareCode(raw)
  const docxKitRuntime = {
    ...docxKit,
    ...overrides,
  }

  // eslint-disable-next-line no-new-func -- sandboxed evaluation is the point of a code playground
  const execute = new Function('docxKit', 'docx', source)

  return execute(docxKitRuntime, docx)
}
