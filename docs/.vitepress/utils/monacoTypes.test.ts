import ts from 'typescript'
import { describe, expect, it } from 'vitest'
import { PRESETS } from '../constants/templates'
import { DOCX_KIT_TYPE_LIBS } from './monacoTypes'

const MAIN_FILE = '/main.ts'

function createTypeService() {
  const files = new Map(
    DOCX_KIT_TYPE_LIBS.map(typeLib => [
      toVirtualFileName(typeLib.filePath),
      typeLib.content,
    ]),
  )
  let mainSource = ''
  let mainVersion = 0

  const service = ts.createLanguageService({
    readDirectory: ts.sys.readDirectory,
    fileExists: fileName => files.has(fileName) || ts.sys.fileExists(fileName),
    getCurrentDirectory: () => '/',
    getDefaultLibFileName: options => ts.getDefaultLibFilePath(options),
    getScriptFileNames: () => [MAIN_FILE, ...files.keys()],
    readFile: fileName => files.get(fileName) ?? ts.sys.readFile(fileName),
    getCompilationSettings: () => ({
      lib: ['lib.esnext.d.ts', 'lib.dom.d.ts'],
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.NodeJs,
      noEmit: true,
      skipLibCheck: true,
      strict: true,
      target: ts.ScriptTarget.ESNext,
    }),
    getScriptSnapshot: fileName => {
      if (fileName === MAIN_FILE) {
        return ts.ScriptSnapshot.fromString(mainSource)
      }

      const source = files.get(fileName) ?? ts.sys.readFile(fileName)
      return source === undefined
        ? undefined
        : ts.ScriptSnapshot.fromString(source)
    },
    getScriptVersion: fileName =>
      fileName === MAIN_FILE ? String(mainVersion) : '0',
  })

  return {
    getDiagnostics(source: string) {
      mainSource = source
      mainVersion++

      return [
        ...service.getSyntacticDiagnostics(MAIN_FILE),
        ...service.getSemanticDiagnostics(MAIN_FILE),
      ]
    },
  }
}

function formatDiagnostics(diagnostics: readonly ts.Diagnostic[]) {
  return diagnostics
    .map(diagnostic => {
      const message = ts.flattenDiagnosticMessageText(
        diagnostic.messageText,
        '\n',
      )
      if (!diagnostic.file || diagnostic.start === undefined) {
        return message
      }

      const position = diagnostic.file.getLineAndCharacterOfPosition(
        diagnostic.start,
      )
      return `${position.line + 1}:${position.character + 1} ${message}`
    })
    .join('\n')
}

function toVirtualFileName(filePath: string) {
  return decodeURIComponent(new URL(filePath).pathname)
}

describe('playground Monaco declarations', () => {
  it('converts file URLs without requiring a Windows drive letter', () => {
    expect(toVirtualFileName('file:///node_modules/docx-kit/index.d.ts')).toBe(
      '/node_modules/docx-kit/index.d.ts',
    )
  })

  it('type-checks every playground preset against the real package API', () => {
    const typeService = createTypeService()

    for (const preset of PRESETS) {
      const diagnostics = typeService.getDiagnostics(preset.code)
      expect(
        formatDiagnostics(diagnostics),
        `${preset.label} contains invalid playground types`,
      ).toBe('')
    }
  })
})
