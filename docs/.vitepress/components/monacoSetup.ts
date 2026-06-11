/**
 * Monaco Editor setup module.
 * This file MUST only be imported dynamically in `onMounted` to
 * avoid SSR issues — Monaco references `window` and imports CSS.
 */

// ─── Vite ?worker type declarations ────────────────────────────────────
// ─── Imports ────────────────────────────────────────────────────────────
import * as monaco from 'monaco-editor'
import EditorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker'
import CSSWorker from 'monaco-editor/esm/vs/language/css/css.worker?worker'
import HTMLWorker from 'monaco-editor/esm/vs/language/html/html.worker?worker'
import JSONWorker from 'monaco-editor/esm/vs/language/json/json.worker?worker'
import TSWorker from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker'

const globalSelf = globalThis as any

// ─── Monaco Environment ─────────────────────────────────────────────────
globalSelf.MonacoEnvironment = {
  getWorker(_workerId: string, label: string): Worker {
    switch (label) {
      case 'css':
      case 'less':
      case 'scss':
        return new CSSWorker()
      case 'handlebars':
      case 'html':
      case 'razor':
        return new HTMLWorker()
      case 'javascript':
      case 'typescript':
        return new TSWorker()
      case 'json':
        return new JSONWorker()
      default:
        return new EditorWorker()
    }
  },
}

export { monaco }
