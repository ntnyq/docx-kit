import type { DocxKitConfig } from 'docx-kit'

export interface PlaygroundRuntimeOverrides {
  transformConfig?: (config: DocxKitConfig) => DocxKitConfig
}

let runtimePromise:
  | Promise<{
      docx: typeof import('docx')
      docxKit: typeof import('docx-kit')
      prepareCode: typeof import('./prepareCode').prepareCode
    }>
  | undefined

/**
 * Execute playground source against the complete browser runtime exports.
 *
 * Keeping `docx-kit` as a namespace makes runtime availability match the
 * declarations exposed to Monaco, including every built-in plugin.
 *
 * Heavy runtime dependencies are loaded on the first Run action instead of
 * being included in the initial playground route chunk.
 */
export async function executePlaygroundCode(
  raw: string,
  overrides: PlaygroundRuntimeOverrides = {},
): Promise<unknown> {
  const { docx, docxKit, prepareCode } = await loadPlaygroundRuntime()
  const source = prepareCode(raw)
  const docxKitRuntime = {
    ...docxKit,
    ...(overrides.transformConfig
      ? {
          createDocx: (config: DocxKitConfig = {}) =>
            docxKit.createDocx(overrides.transformConfig!(config)),
        }
      : {}),
  }

  // eslint-disable-next-line no-new-func -- sandboxed evaluation is the point of a code playground
  const execute = new Function('docxKit', 'docx', source)

  return execute(docxKitRuntime, docx)
}

function loadPlaygroundRuntime() {
  runtimePromise ??= Promise.all([
    import('docx'),
    import('docx-kit'),
    import('./prepareCode'),
  ]).then(([docx, docxKit, { prepareCode }]) => ({
    docx,
    docxKit,
    prepareCode,
  }))

  return runtimePromise
}
