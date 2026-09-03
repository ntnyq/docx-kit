/**
 * Monaco Editor setup module.
 * This file MUST only be imported dynamically in `onMounted` to
 * avoid SSR issues — Monaco references `window` and imports CSS.
 */

// ─── Imports ────────────────────────────────────────────────────────────
import * as monaco from 'monaco-editor/editor/editor.api'
import EditorWorker from 'monaco-editor/editor/editor.worker?worker'
import TSWorker from 'monaco-editor/language/typescript/ts.worker?worker'
import * as typescript from 'monaco-editor/languages/features/typescript/register'

// Syntax highlighting is registered separately from TypeScript language services.
import 'monaco-editor/languages/definitions/typescript/register'

const globalSelf = globalThis as {
  MonacoEnvironment?: {
    getWorker(_workerId: string, label: string): Worker
  }
} & typeof globalThis

// ─── Monaco Environment ─────────────────────────────────────────────────
globalSelf.MonacoEnvironment = {
  getWorker(_workerId: string, label: string): Worker {
    switch (label) {
      case 'javascript':
      case 'typescript':
        return new TSWorker()
      default:
        return new EditorWorker()
    }
  },
}

export { monaco, typescript }
